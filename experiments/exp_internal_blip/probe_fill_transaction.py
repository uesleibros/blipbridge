"""Observe the loaded-image transfer, the OART operation lifecycle and refcounts.

Run inside GDB after ``tools/prepare_decoder_trace.ps1``; see
``tools/run_office_probe.ps1`` for the supported invocation. This
experiment only reads registers and bounded live records. It never calls a
private Office function, never writes inferior memory, and never retains an
Office pointer after detaching.

All addresses are observational on OART 16.0.14334.20848 x64 and GFX
16.0.14334.20848 x64. Every breakpoint address is validated against recorded
instruction bytes before it is set; a mismatch aborts the run.
"""

import json
import pathlib
import struct

import gdb


OFFICE_BUILD = '16.0.14334.20848'
POINTER_SIZE = 8

# ---------------------------------------------------------------------------
# Stage 1: UserPicture loads the file, then transfers the loaded image record.
# ---------------------------------------------------------------------------
# UserPicture calls this wrapper after the filename loader returns. It copies
# the loaded record into destination+8, then changes low discriminator bits.
# Evidence: docs/resource_lifetime.md and OART +0x89C9F6 disassembly.
RECORD_TRANSFER_RVA = 0x22BCF4
RECORD_TRANSFER_SIGNATURE = bytes.fromhex('40 53 48 83 ec 20 48 8b d9')

# Offset of the cached GFX image pointer inside the image sub-record. The same
# offset appears in the standalone loaded record and, at sub-record base +0x90,
# inside the enclosing property record.
CACHED_RESOURCE_OFFSET = 0xF0

# GFX intrusive objects keep a 32-bit reference count at object+8. Release is
# vtable slot +8: GFX +0xA1A0 runs `lock xadd` on +8 and destroys at 1.
# Evidence: docs/resource_lifetime.md and GFX +0xA1A0 disassembly.
GFX_REFERENCE_COUNT_OFFSET = 8
GFX_REFERENCE_COUNT_SIZE = 4

# ---------------------------------------------------------------------------
# Stage 2: the transaction is constructed and handed to the OART receiver.
# ---------------------------------------------------------------------------
# The subsequent call goes through the internal receiver's vtable slot +0x78.
# RAX holds the resolved target, RCX the receiver, RDX the transaction record.
TRANSACTION_CALL_RVA = 0x89CA6B
TRANSACTION_CALL_SIGNATURE = bytes.fromhex('ff 15 57 9f 16 00')
TRANSACTION_RETURN_RVA = 0x89CA71
TRANSACTION_RETURN_SIGNATURE = bytes.fromhex('48 8b 4c 24 38')

# Only inspect the first 0x100 bytes of an image record. Existing static code
# accesses it beyond +0x1E8; never scan heap regions for matching pointers.
RECORD_OBSERVATION_SIZE = 0x100
DESTINATION_RECORD_OFFSET = 8
DESTINATION_OBSERVATION_SIZE = RECORD_OBSERVATION_SIZE + DESTINATION_RECORD_OFFSET
# Constructor OART +0x48870 writes through transaction+0x508. The caller
# reserves this region below its stack cookie. This bound applies only to that
# record; the operation factory reads transaction+0x500..+0x508 as scalars.
TRANSACTION_OBSERVATION_SIZE = 0x510
# The receiver is resolved by OART +0x63EA0 from handler state +0x58. Its
# vtable is OART +0x9F6658; slot +0xF0 returns the context sub-object.
RECEIVER_OBSERVATION_SIZE = 0x40
# OART +0x21DEC0 forwards to receiver vtable+0x50 without adjusting arguments.
FORWARDED_METHOD_SLOT = 0x50

# ---------------------------------------------------------------------------
# Stage 3: build -> apply -> delete, driven by OART +0x21BB20.
# ---------------------------------------------------------------------------
# +0x21BB20 calls receiver vtable+0x60 to build the operation, vtable+0x58 to
# apply it, then operation vtable+0xA8 with EDX=1 to destroy and free it.
BUILD_CALL_RVA = 0x21BBE8
BUILD_CALL_SIGNATURE = bytes.fromhex('ff 15 da ad 7e 00')
APPLY_CALL_RVA = 0x21BC06
APPLY_CALL_SIGNATURE = bytes.fromhex('ff 15 bc ad 7e 00')
DELETE_CALL_RVA = 0x21BC26
DELETE_CALL_SIGNATURE = bytes.fromhex('ff 15 9c ad 7e 00')

# Operation allocation inside factory OART +0xE7F0. EDX carries the requested
# size; the allocator singleton lives at OART +0xD40050 and allocates through
# its vtable slot 0.
OPERATION_ALLOCATION_CALL_RVA = 0xEF0E
OPERATION_ALLOCATION_CALL_SIGNATURE = bytes.fromhex('ff 15 b4 7a 9f 00')
# Static allocation size for the branch that builds the +0x9E4BD0 operation.
EXPECTED_OPERATION_SIZE = 0x570

# The call that constructs the operation. Arguments at this instruction:
#   RCX  = freshly allocated operation storage
#   RDX  = transaction+0x18 (the property record copied into the operation)
#   R8D  = transaction+0x500 (flags)
#   R9B  = transaction+0x508 (bool)
#   [RSP+0x20] = transaction+0x504 (property/type identifier)
OPERATION_FACTORY_CALL_RVA = 0xEF3B
OPERATION_FACTORY_CALL_SIGNATURE = bytes.fromhex('e8 04 13 00 00')
OPERATION_FACTORY_STACK_ARGUMENT_OFFSET = 0x20

# Constructor entry, reached only from the factory call above.
OPERATION_CONSTRUCTOR_RVA = 0x10244
OPERATION_CONSTRUCTOR_SIGNATURE = bytes.fromhex('48 89 5c 24 18 55 56 57')
# The constructor embeds its copy of the property record at operation+0x68, so
# the cached image lands at 0x68 + 0x90 + 0xF0.
OPERATION_RECORD_OFFSET = 0x68
IMAGE_SUB_RECORD_OFFSET = 0x90
OPERATION_CACHED_IMAGE_OFFSET = (
    OPERATION_RECORD_OFFSET + IMAGE_SUB_RECORD_OFFSET + CACHED_RESOURCE_OFFSET)

# Application helper invoked by OART +0x1B88B0 with the operation, the context
# returned by receiver vtable+0xF0, and receiver+0x10.
OPERATION_EXECUTION_RVA = 0x1B8F50
OPERATION_EXECUTION_SIGNATURE = bytes.fromhex('48 8b c4 4c 89 48 20')

# Scalar deleting destructor OART +0x66C80 destroys the embedded record at
# +0x68 first, then frees the object through allocator slot +8 when EDX bit 0
# is set. Sampling here reports the count after the operation released its own
# reference. Never read the operation after this call returns.
OPERATION_DEALLOCATION_CALL_RVA = 0x66CDE
OPERATION_DEALLOCATION_CALL_SIGNATURE = bytes.fromhex('ff 15 e4 fc 99 00')


class TransactionProbe:
    """Owns temporary breakpoints; all recorded addresses are borrowed.

    No field of this class outlives :meth:`execute`. The probe detaches in a
    ``finally`` block so a failed observation still leaves Office running
    normally and without breakpoints.
    """

    def __init__(self):
        self.root = pathlib.Path.cwd()
        self.target = json.loads(
            (self.root / 'artifacts/decoder_target.json').read_text(encoding='utf-8-sig'))
        self.breakpoints = []
        self.cached_resource = None
        self.thread_id = None
        self.destination = None
        self.transaction = None
        self.operation = None
        self.counts = []

    # -- primitive reads ---------------------------------------------------
    def read(self, address, size):
        return bytes(gdb.selected_inferior().read_memory(address, size))

    def pointer(self, address):
        return struct.unpack('<Q', self.read(address, POINTER_SIZE))[0]

    def register(self, name):
        return int(gdb.parse_and_eval('$' + name)) & 0xFFFFFFFFFFFFFFFF

    def location(self, address):
        for module in self.target['modules']:
            if module['base'] <= address < module['base'] + module['size']:
                return '%s+0x%x' % (module['name'], address - module['base'])
        return hex(address)

    # -- reference-count sampling -----------------------------------------
    def sample_count(self, label):
        """Record the cached image's intrusive count without modifying it."""
        if self.cached_resource is None:
            return None
        raw = self.read(self.cached_resource + GFX_REFERENCE_COUNT_OFFSET,
                        GFX_REFERENCE_COUNT_SIZE)
        count = struct.unpack('<I', raw)[0]
        self.counts.append((label, count))
        gdb.write('CACHED COUNT %-22s = %d\n' % (label, count))
        return count

    def record(self, label, address, size=RECORD_OBSERVATION_SIZE):
        gdb.write('%s address=%#x\n' % (label, address))
        data = self.read(address, size)
        for offset in range(0, len(data), POINTER_SIZE):
            value = struct.unpack_from('<Q', data, offset)[0]
            if value:
                gdb.write('  +%#x = %s%s\n' % (
                    offset, self.location(value),
                    ' [cached resource]' if value == self.cached_resource else ''))

    # -- driver ------------------------------------------------------------
    def execute(self):
        gdb.execute('set pagination off')
        gdb.execute('set confirm off')
        gdb.execute('set solib-search-path C:/Program Files/Microsoft Office/root/vfs/'
                    'ProgramFilesCommonX64/Microsoft Shared/Office16;'
                    'C:/Program Files/Microsoft Office/root/Office16')
        gdb.execute('attach ' + str(self.target['pid']))
        try:
            oart = next(module for module in self.target['modules']
                        if module['name'].lower() == 'oart.dll')
            if oart.get('version') != OFFICE_BUILD:
                raise RuntimeError('Unsupported OART build')
            profiles = [
                (RECORD_TRANSFER_RVA, RECORD_TRANSFER_SIGNATURE, 'transfer'),
                (TRANSACTION_CALL_RVA, TRANSACTION_CALL_SIGNATURE, 'transaction'),
                (TRANSACTION_RETURN_RVA, TRANSACTION_RETURN_SIGNATURE, 'complete'),
                (BUILD_CALL_RVA, BUILD_CALL_SIGNATURE, 'build'),
                (APPLY_CALL_RVA, APPLY_CALL_SIGNATURE, 'apply'),
                (DELETE_CALL_RVA, DELETE_CALL_SIGNATURE, 'delete'),
                (OPERATION_ALLOCATION_CALL_RVA, OPERATION_ALLOCATION_CALL_SIGNATURE,
                 'allocate'),
                (OPERATION_FACTORY_CALL_RVA, OPERATION_FACTORY_CALL_SIGNATURE, 'factory'),
                (OPERATION_CONSTRUCTOR_RVA, OPERATION_CONSTRUCTOR_SIGNATURE,
                 'constructor'),
                (OPERATION_EXECUTION_RVA, OPERATION_EXECUTION_SIGNATURE, 'execution'),
                (OPERATION_DEALLOCATION_CALL_RVA, OPERATION_DEALLOCATION_CALL_SIGNATURE,
                 'deallocate'),
            ]
            for rva, signature, stage in profiles:
                address = oart['base'] + rva
                if self.read(address, len(signature)) != signature:
                    raise RuntimeError('Signature mismatch at ' + hex(rva))
                self.breakpoints.append(TransactionBreakpoint(address, stage, self))
            (self.root / 'artifacts/decoder_go').write_text('go')
            gdb.execute('continue')
        finally:
            self.summarize()
            for breakpoint in self.breakpoints:
                breakpoint.delete()
            gdb.execute('detach')

    def summarize(self):
        """Print the count timeline and the two deltas that decide ownership."""
        if not self.counts:
            return
        gdb.write('\nCACHED REFERENCE COUNT TIMELINE\n')
        for label, count in self.counts:
            gdb.write('  %-24s %d\n' % (label, count))
        sampled = dict(self.counts)
        before = sampled.get('before-construction')
        after = sampled.get('after-construction')
        held = sampled.get('before-destruction')
        released = sampled.get('after-destruction')
        if before is not None and after is not None:
            gdb.write('  construction delta = %+d\n' % (after - before))
        if held is not None and released is not None:
            gdb.write('  destruction delta  = %+d\n' % (released - held))


class TransactionBreakpoint(gdb.Breakpoint):
    """One observation point. ``stop`` never resumes with modified state."""

    def __init__(self, address, stage, probe):
        super().__init__('*' + hex(address), internal=True)
        self.stage = stage
        self.probe = probe

    # Each handler returns False so the inferior continues; the probe never
    # leaves Office stopped inside an internal Office frame.
    def stop(self):
        probe = self.probe
        current_thread = gdb.selected_thread().ptid
        if probe.thread_id is not None and current_thread != probe.thread_id:
            return False
        handler = getattr(self, 'on_' + self.stage.replace('-', '_'))
        return handler(probe)

    # -- stage 1 -----------------------------------------------------------
    def on_transfer(self, probe):
        probe.thread_id = gdb.selected_thread().ptid
        source = probe.register('rdx')
        probe.cached_resource = probe.pointer(source + CACHED_RESOURCE_OFFSET)
        probe.destination = probe.register('rcx')
        gdb.write('\nSTAGE transfer\n')
        probe.record('loaded record', source)
        probe.record('destination before transfer', probe.destination)
        probe.sample_count('loaded-record')
        self.enabled = False
        return False

    # -- stage 2 -----------------------------------------------------------
    def on_transaction(self, probe):
        probe.transaction = probe.register('rdx')
        receiver = probe.register('rcx')
        gdb.write('\nSTAGE transaction\n')
        gdb.write('receiver=%#x vtable=%s resolved-call=%s\n' % (
            receiver, probe.location(probe.pointer(receiver)),
            probe.location(probe.register('rax'))))
        forwarded = probe.pointer(probe.pointer(receiver) + FORWARDED_METHOD_SLOT)
        gdb.write('Forwarded slot +0x50 target=%s\n' % probe.location(forwarded))
        probe.record('receiver', receiver, RECEIVER_OBSERVATION_SIZE)
        # Destination embeds the loaded image record at +8; include its
        # complete cached pointer at +0xF8 in the bounded observation.
        probe.record('destination after transfer', probe.destination,
                     DESTINATION_OBSERVATION_SIZE)
        probe.record('transaction', probe.transaction, TRANSACTION_OBSERVATION_SIZE)
        probe.sample_count('transaction-entry')
        gdb.execute('bt 12')
        self.enabled = False
        return False

    def on_complete(self, probe):
        if probe.transaction is None:
            return False
        gdb.write('\nSTAGE complete\n')
        probe.sample_count('transaction-return')
        gdb.write('Transaction returned; RAX=%#x (return type not established)\n' %
                  probe.register('rax'))
        return True

    # -- stage 3 -----------------------------------------------------------
    def on_build(self, probe):
        if probe.transaction is None:
            return False
        gdb.write('\nSTAGE build target=%s receiver=%#x transaction=%#x\n' % (
            probe.location(probe.register('rax')), probe.register('rcx'),
            probe.register('rdx')))
        return False

    def on_allocate(self, probe):
        if probe.transaction is None:
            return False
        size = probe.register('rdx') & 0xFFFFFFFF
        gdb.write('\nOPERATION allocate size=%#x expected=%#x match=%s\n' % (
            size, EXPECTED_OPERATION_SIZE, size == EXPECTED_OPERATION_SIZE))
        return False

    def on_factory(self, probe):
        if probe.transaction is None:
            return False
        stack_argument = probe.pointer(
            probe.register('rsp') + OPERATION_FACTORY_STACK_ARGUMENT_OFFSET) & 0xFFFFFFFF
        source = probe.register('rdx')
        gdb.write('\nOPERATION factory storage=%#x source=%#x flags=%#x '
                  'bool=%#x identifier=%#x\n' % (
                      probe.register('rcx'), source,
                      probe.register('r8') & 0xFFFFFFFF,
                      probe.register('r9') & 0xFF, stack_argument))
        gdb.write('Source equals transaction+0x18: %s\n' %
                  (source == probe.transaction + 0x18))
        probe.sample_count('before-construction')
        gdb.execute('bt 8')
        return False

    def on_constructor(self, probe):
        if probe.transaction is None:
            return False
        probe.operation = probe.register('rcx')
        gdb.write('\nOPERATION constructor entered this=%#x\n' % probe.operation)
        return False

    def on_apply(self, probe):
        if probe.operation is None:
            return False
        operation = probe.register('rdx')
        cached = probe.pointer(operation + OPERATION_CACHED_IMAGE_OFFSET)
        gdb.write('\nSTAGE apply target=%s receiver=%#x operation=%#x\n' % (
            probe.location(probe.register('rax')), probe.register('rcx'), operation))
        gdb.write('Operation is the constructed object: %s\n' %
                  (operation == probe.operation))
        gdb.write('Operation vtable=%s\n' % probe.location(probe.pointer(operation)))
        gdb.write('Operation cached image +%#x = %#x equals tracked GFX: %s\n' % (
            OPERATION_CACHED_IMAGE_OFFSET, cached, cached == probe.cached_resource))
        probe.sample_count('after-construction')
        return False

    def on_execution(self, probe):
        if probe.operation is None:
            return False
        gdb.write('\nOPERATION execution operation=%#x context=%#x receiver_field=%#x '
                  'holder=%#x\n' % (probe.register('rcx'), probe.register('rdx'),
                                    probe.register('r8'), probe.register('r9')))
        gdb.write('Execution RCX equals constructed operation: %s\n' %
                  (probe.register('rcx') == probe.operation))
        gdb.execute('bt 8')
        return False

    def on_delete(self, probe):
        if probe.operation is None:
            return False
        gdb.write('\nSTAGE delete target=%s operation=%#x flags=%#x\n' % (
            probe.location(probe.register('rax')), probe.register('rcx'),
            probe.register('rdx') & 0xFFFFFFFF))
        gdb.write('Delete target is the constructed operation: %s\n' %
                  (probe.register('rcx') == probe.operation))
        probe.sample_count('before-destruction')
        return False

    def on_deallocate(self, probe):
        if probe.operation is None or probe.register('rdx') != probe.operation:
            return False
        gdb.write('\nOPERATION allocator-release address=%#x target=%s\n' % (
            probe.operation, probe.location(probe.register('rax'))))
        # The embedded record was destroyed before this call, so the count now
        # excludes the operation's own reference.
        probe.sample_count('after-destruction')
        return False


TransactionProbe().execute()
