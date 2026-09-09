"""Observe the loaded-image transfer and subsequent Office state transaction.

Run inside GDB after tools/prepare_decoder_trace.ps1. This experiment reads
registers and bounded live records; it never calls a private Office function.
All addresses below are observational on OART 16.0.14334.20848 x64.
"""

import json
import pathlib
import struct

import gdb


OFFICE_BUILD = '16.0.14334.20848'
POINTER_SIZE = 8

# UserPicture calls this wrapper after the filename loader returns. It copies
# the loaded record into destination+8, then changes low discriminator bits.
# Evidence: docs/resource_lifetime.md and OART +0x89C9F6 disassembly.
RECORD_TRANSFER_RVA = 0x22BCF4
RECORD_TRANSFER_SIGNATURE = bytes.fromhex('40 53 48 83 ec 20 48 8b d9')
CACHED_RESOURCE_OFFSET = 0xF0

# The subsequent call goes through the internal receiver's vtable slot +0x78.
# RAX holds the resolved target, RCX receiver, RDX transaction record. This is
# a candidate state-application boundary, not a validated Shape setter ABI.
TRANSACTION_CALL_RVA = 0x89CA6B
TRANSACTION_CALL_SIGNATURE = bytes.fromhex('ff 15 57 9f 16 00')
TRANSACTION_RETURN_RVA = 0x89CA71
TRANSACTION_RETURN_SIGNATURE = bytes.fromhex('48 8b 4c 24 38')

# Only inspect the first 0x100 bytes. Existing static code accesses the loaded
# record beyond +0x1E8 and allocates a larger transaction frame; never scan heap
# regions for matching pointers or dereference arbitrary members.
RECORD_OBSERVATION_SIZE = 0x100
DESTINATION_RECORD_OFFSET = 8
DESTINATION_OBSERVATION_SIZE = RECORD_OBSERVATION_SIZE + DESTINATION_RECORD_OFFSET
# Constructor OART +0x48870 writes through transaction+0x508. Caller reserves
# this region below its stack cookie. This bound applies only to that record.
TRANSACTION_OBSERVATION_SIZE = 0x510
# OART +0x21DEC0 forwards to receiver vtable+0x50 without adjusting arguments.
FORWARDED_METHOD_SLOT = 0x50
# OART +0x21BB20 invokes slot +0x60 with transaction and output storage, then
# slot +0x58 with that output if successful. Semantic aliases remain candidates.
PREPARE_CALL_RVA = 0x21BBE8
PREPARE_CALL_SIGNATURE = bytes.fromhex('ff 15 da ad 7e 00')
COMMIT_CALL_RVA = 0x21BC06
COMMIT_CALL_SIGNATURE = bytes.fromhex('ff 15 bc ad 7e 00')


class TransactionProbe:
    """Owns temporary breakpoints; all recorded addresses are borrowed."""

    def __init__(self):
        self.root = pathlib.Path.cwd()
        self.target = json.loads(
            (self.root / 'artifacts/decoder_target.json').read_text(encoding='utf-8-sig'))
        self.breakpoints = []
        self.cached_resource = None
        self.thread_id = None
        self.destination = None

    def read(self, address, size):
        return bytes(gdb.selected_inferior().read_memory(address, size))

    def pointer(self, address):
        return struct.unpack('<Q', self.read(address, POINTER_SIZE))[0]

    def register(self, name):
        return int(gdb.parse_and_eval('$' + name))

    def location(self, address):
        for module in self.target['modules']:
            if module['base'] <= address < module['base'] + module['size']:
                return '%s+0x%x' % (module['name'], address - module['base'])
        return hex(address)

    def record(self, label, address, size=RECORD_OBSERVATION_SIZE):
        gdb.write('%s address=%#x\n' % (label, address))
        data = self.read(address, size)
        for offset in range(0, len(data), POINTER_SIZE):
            value = struct.unpack_from('<Q', data, offset)[0]
            if value:
                gdb.write('  +%#x = %s%s\n' % (
                    offset, self.location(value),
                    ' [cached resource]' if value == self.cached_resource else ''))

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
                (PREPARE_CALL_RVA, PREPARE_CALL_SIGNATURE, 'prepare-candidate'),
                (COMMIT_CALL_RVA, COMMIT_CALL_SIGNATURE, 'commit-candidate'),
            ]
            for rva, signature, stage in profiles:
                address = oart['base'] + rva
                if self.read(address, len(signature)) != signature:
                    raise RuntimeError('Signature mismatch at ' + hex(rva))
                self.breakpoints.append(TransactionBreakpoint(address, stage, self))
            (self.root / 'artifacts/decoder_go').write_text('go')
            gdb.execute('continue')
        finally:
            for breakpoint in self.breakpoints:
                breakpoint.delete()
            gdb.execute('detach')


class TransactionBreakpoint(gdb.Breakpoint):
    def __init__(self, address, stage, probe):
        super().__init__('*' + hex(address), internal=True)
        self.stage = stage
        self.probe = probe

    def stop(self):
        probe = self.probe
        current_thread = gdb.selected_thread().ptid
        if probe.thread_id is not None and current_thread != probe.thread_id:
            return False
        if self.stage in ['prepare-candidate', 'commit-candidate']:
            if probe.thread_id is None:
                return False
            gdb.write('\nSTAGE %s target=%s receiver=%#x argument=%#x\n' % (
                self.stage, probe.location(probe.register('rax')),
                probe.register('rcx'), probe.register('rdx')))
            if self.stage == 'commit-candidate':
                gdb.write('Output object first qword=%s\n' %
                          probe.location(probe.pointer(probe.register('rdx'))))
            return False
        gdb.write('\nSTAGE ' + self.stage + '\n')
        if self.stage == 'transfer':
            probe.thread_id = current_thread
            source = probe.register('rdx')
            probe.cached_resource = probe.pointer(source + CACHED_RESOURCE_OFFSET)
            probe.destination = probe.register('rcx')
            probe.record('loaded record', source)
            probe.record('destination before transfer', probe.register('rcx'))
            self.enabled = False
            return False
        if self.stage == 'transaction':
            receiver = probe.register('rcx')
            gdb.write('receiver=%#x vtable=%s resolved-call=%s\n' % (
                receiver, probe.location(probe.pointer(receiver)),
                probe.location(probe.register('rax'))))
            forwarded = probe.pointer(probe.pointer(receiver) + FORWARDED_METHOD_SLOT)
            gdb.write('Forwarded slot +0x50 target=%s\n' % probe.location(forwarded))
            # Destination embeds the loaded image record at +8; include its
            # complete cached pointer at +0xF8 in the bounded observation.
            probe.record('destination after transfer', probe.destination,
                         DESTINATION_OBSERVATION_SIZE)
            probe.record('transaction', probe.register('rdx'), TRANSACTION_OBSERVATION_SIZE)
            gdb.execute('bt 12')
            self.enabled = False
            return False
        gdb.write('Transaction returned; RAX=%#x (return type not established)\n' %
                  probe.register('rax'))
        return True


TransactionProbe().execute()
