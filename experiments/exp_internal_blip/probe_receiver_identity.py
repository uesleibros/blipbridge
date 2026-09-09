"""Compare the OART fill receiver across several Shapes in one presentation.

Run inside GDB after ``tools/prepare_receiver_identity.ps1``; see
``tools/run_office_probe.ps1`` for the supported invocation. Three ordinary
AutoShapes - two on one slide, one on another - receive the same picture, so the
only thing that varies between the observed calls is Shape and slide context.

The question this answers is narrow: is the receiver passed to the fill
transaction per Shape, per slide or per document, and which of its fields (or of
the property record below the image sub-record) change with the target.

This experiment only reads registers and bounded records. It never calls a
private Office function, never writes inferior memory, and never retains an
Office pointer after detaching. All addresses are observational on OART
16.0.14334.20848 x64 and are byte-validated before any breakpoint is set.
"""

import json
import pathlib
import struct

import gdb


OFFICE_BUILD = '16.0.14334.20848'
POINTER_SIZE = 8

# `mov 0x58(%r15),%rcx` feeds OART +0x63EA0, which resolves the receiver from a
# token held by the UserPicture handler state. At the call, R15 is that handler
# state and RCX is the token. Evidence: OART +0x89C860 disassembly.
RECEIVER_RESOLVE_CALL_RVA = 0x89CA52
RECEIVER_RESOLVE_CALL_SIGNATURE = bytes.fromhex('e8 49 74 7c ff')

# The transaction is handed to the receiver through its vtable slot +0x78.
TRANSACTION_CALL_RVA = 0x89CA6B
TRANSACTION_CALL_SIGNATURE = bytes.fromhex('ff 15 57 9f 16 00')

# Discriminators for the picture-fill transaction observed in
# docs/fill_transaction.md. A hit that fails either check is ignored rather than
# guessed at.
TRANSACTION_VTABLE_RVA = 0x9ED7E0
RECEIVER_VTABLE_RVA = 0x9F6658

# Layout constants recovered in docs/fill_transaction.md.
PROPERTY_RECORD_OFFSET = 0x18
IMAGE_SUB_RECORD_OFFSET = 0x90
CACHED_RESOURCE_OFFSET = 0xF0
TRANSACTION_CACHED_IMAGE_OFFSET = (
    PROPERTY_RECORD_OFFSET + IMAGE_SUB_RECORD_OFFSET + CACHED_RESOURCE_OFFSET)

# Bounded observations. The record prefix is exactly the part below the image
# sub-record, which is the region whose contents have never been mapped.
RECORD_PREFIX_SIZE = IMAGE_SUB_RECORD_OFFSET
RECEIVER_OBSERVATION_SIZE = 0x40

# Two per-call pointers worth typing rather than only diffing: the object the
# receiver holds at +0x8, and the only pointer in the record prefix.
RECEIVER_CONTAINER_OFFSET = 0x8
RECORD_REFERENCE_OFFSET = 0x40

EXPECTED_OBSERVATIONS = 3


class ReceiverProbe:
    """Owns temporary breakpoints; all recorded addresses are borrowed."""

    def __init__(self):
        self.root = pathlib.Path.cwd()
        self.target = json.loads(
            (self.root / 'artifacts/decoder_target.json').read_text(encoding='utf-8-sig'))
        self.oart = next(module for module in self.target['modules']
                         if module['name'].lower() == 'oart.dll')
        self.breakpoints = []
        self.thread_id = None
        self.pending_handler = None
        self.pending_token = None
        self.observations = []

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

    def rva(self, rva):
        return self.oart['base'] + rva

    def described(self, address):
        """Describe the pointer at ``address`` by target address and vtable."""
        value = self.pointer(address)
        if not value:
            return 'null'
        try:
            return '%#x vtable=%s' % (value, self.location(self.pointer(value)))
        except gdb.MemoryError:
            return '%#x (unreadable)' % value

    def words(self, address, size):
        data = self.read(address, size)
        return [struct.unpack_from('<Q', data, offset)[0]
                for offset in range(0, size, POINTER_SIZE)]

    # -- driver ------------------------------------------------------------
    def execute(self):
        gdb.execute('set pagination off')
        gdb.execute('set confirm off')
        gdb.execute('set solib-search-path C:/Program Files/Microsoft Office/root/vfs/'
                    'ProgramFilesCommonX64/Microsoft Shared/Office16;'
                    'C:/Program Files/Microsoft Office/root/Office16')
        gdb.execute('attach ' + str(self.target['pid']))
        try:
            if self.oart.get('version') != OFFICE_BUILD:
                raise RuntimeError('Unsupported OART build')
            profiles = [
                (RECEIVER_RESOLVE_CALL_RVA, RECEIVER_RESOLVE_CALL_SIGNATURE, 'resolve'),
                (TRANSACTION_CALL_RVA, TRANSACTION_CALL_SIGNATURE, 'transaction'),
            ]
            for rva, signature, stage in profiles:
                address = self.rva(rva)
                if self.read(address, len(signature)) != signature:
                    raise RuntimeError('Signature mismatch at ' + hex(rva))
                self.breakpoints.append(ReceiverBreakpoint(address, stage, self))
            (self.root / 'artifacts/decoder_go').write_text('go')
            gdb.execute('continue')
        finally:
            self.summarize()
            for breakpoint in self.breakpoints:
                breakpoint.delete()
            # PowerPoint quits once the last presentation closes, so the
            # inferior may already be gone by the time the probe unwinds.
            try:
                gdb.execute('detach')
            except gdb.error:
                gdb.write('Inferior already exited; nothing to detach.\n')

    # -- reporting ---------------------------------------------------------
    def record_observation(self, observation):
        index = len(self.observations)
        self.observations.append(observation)
        gdb.write('\nOBSERVATION %d\n' % index)
        gdb.write('  handler state   = %#x\n' % observation['handler'])
        gdb.write('  receiver token  = %#x\n' % observation['token'])
        gdb.write('  receiver        = %#x\n' % observation['receiver'])
        gdb.write('  cached image    = %#x\n' % observation['cached'])
        gdb.write('  receiver +%#x   = %s\n' % (
            RECEIVER_CONTAINER_OFFSET, observation['container']))
        gdb.write('  record   +%#x   = %s\n' % (
            RECORD_REFERENCE_OFFSET, observation['record_reference']))
        for offset, value in enumerate(observation['receiver_words']):
            if value:
                gdb.write('  receiver +%#x = %s\n' % (
                    offset * POINTER_SIZE, self.location(value)))
        for offset, value in enumerate(observation['record_words']):
            if value:
                gdb.write('  record   +%#x = %s\n' % (
                    offset * POINTER_SIZE, self.location(value)))
        if len(self.observations) >= EXPECTED_OBSERVATIONS:
            gdb.write('\nCollected %d observations; releasing breakpoints.\n' %
                      len(self.observations))
            for breakpoint in self.breakpoints:
                breakpoint.enabled = False

    def summarize(self):
        """Report what varied between Shapes; say nothing about what did not."""
        if len(self.observations) < 2:
            gdb.write('\nToo few observations to compare.\n')
            return
        gdb.write('\nRECEIVER COMPARISON ACROSS %d SHAPES\n' % len(self.observations))
        for name in ['handler', 'token', 'receiver', 'cached']:
            values = [observation[name] for observation in self.observations]
            gdb.write('  %-16s %s -> %s\n' % (
                name, ' '.join('%#x' % value for value in values),
                'identical' if len(set(values)) == 1 else 'differs per call'))
        for name in ['container', 'record_reference']:
            values = [observation[name] for observation in self.observations]
            gdb.write('  %-16s %s\n' % (
                name, 'identical' if len(set(values)) == 1 else 'differs per call'))
            for index, value in enumerate(values):
                gdb.write('      call %d: %s\n' % (index, value))

        for label, key, size in [('receiver', 'receiver_words', RECEIVER_OBSERVATION_SIZE),
                                 ('record', 'record_words', RECORD_PREFIX_SIZE)]:
            differing = []
            for index in range(size // POINTER_SIZE):
                values = [observation[key][index] for observation in self.observations]
                if len(set(values)) != 1:
                    differing.append((index * POINTER_SIZE, values))
            if not differing:
                gdb.write('  %s words: identical at every observed offset\n' % label)
                continue
            gdb.write('  %s words differing between calls:\n' % label)
            for offset, values in differing:
                gdb.write('    +%#x : %s\n' % (
                    offset, ' '.join('%#x' % value for value in values)))


class ReceiverBreakpoint(gdb.Breakpoint):
    """One observation point. ``stop`` always resumes the inferior."""

    def __init__(self, address, stage, probe):
        super().__init__('*' + hex(address), internal=True)
        self.stage = stage
        self.probe = probe

    def stop(self):
        probe = self.probe
        current_thread = gdb.selected_thread().ptid
        if probe.thread_id is not None and current_thread != probe.thread_id:
            return False
        if self.stage == 'resolve':
            # Remember the handler state and token; the receiver itself is only
            # known once OART +0x63EA0 has returned.
            probe.pending_handler = probe.register('r15')
            probe.pending_token = probe.register('rcx')
            return False

        receiver = probe.register('rcx')
        transaction = probe.register('rdx')
        if probe.pointer(receiver) != probe.rva(RECEIVER_VTABLE_RVA):
            return False
        if probe.pointer(transaction) != probe.rva(TRANSACTION_VTABLE_RVA):
            return False
        cached = probe.pointer(transaction + TRANSACTION_CACHED_IMAGE_OFFSET)
        if not cached:
            return False
        probe.thread_id = current_thread
        probe.record_observation({
            'handler': probe.pending_handler or 0,
            'token': probe.pending_token or 0,
            'receiver': receiver,
            'cached': cached,
            'receiver_words': probe.words(receiver, RECEIVER_OBSERVATION_SIZE),
            'record_words': probe.words(transaction + PROPERTY_RECORD_OFFSET,
                                        RECORD_PREFIX_SIZE),
            'container': probe.described(receiver + RECEIVER_CONTAINER_OFFSET),
            'record_reference': probe.described(
                transaction + PROPERTY_RECORD_OFFSET + RECORD_REFERENCE_OFFSET),
        })
        return False


ReceiverProbe().execute()
