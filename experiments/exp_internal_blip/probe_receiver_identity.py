"""Compare the OART fill receiver across several Shapes in one presentation.

Run inside GDB after ``tools/prepare_receiver_identity.ps1``; see
``tools/run_office_probe.ps1`` for the supported invocation. Three ordinary
AutoShapes - two on one slide, one on another - receive the same picture in the
order A, A, B, C, so the only thing that varies between the observed calls is
Shape and slide context.

The repeat on shape A is what makes the classification sound: a value that
changes between calls 0 and 1 cannot denote the Shape, however neatly it happens
to differ between different Shapes. The question this answers is whether the
receiver handed to the fill transaction is per call, per Shape, per slide or per
document, and which of its fields - or of the property record below the image
sub-record - track the target.

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
# Written by receiver constructor OART +0x223950 as a process-global sequence
# number from [OART +0xD40038]; it is an allocation counter, not an identifier
# of anything in the document.
RECEIVER_SEQUENCE_OFFSET = 0x20
# The token at handler+0x58 is a control block: strong count at +0, weak count
# at +4, pointee at +0x10. Handler factory OART +0x23B220 AddRefs it.
TOKEN_STRONG_COUNT_OFFSET = 0x0
TOKEN_POINTEE_OFFSET = 0x10
# Enough of an unknown target to recognise a repeat, without scanning heap.
REFERENCE_TARGET_SIZE = 0x30

# Call order in tools/prepare_receiver_identity.ps1.
CALL_LABELS = ['shape A', 'shape A again', 'shape B same slide', 'shape C other slide']
EXPECTED_OBSERVATIONS = len(CALL_LABELS)


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

    def dword(self, address):
        return struct.unpack('<I', self.read(address, 4))[0]

    def snapshot(self, address):
        """Hex prefix of an unknown target, only to recognise repeats."""
        if not address:
            return 'null'
        try:
            return self.read(address, REFERENCE_TARGET_SIZE).hex()
        except gdb.MemoryError:
            return 'unreadable'

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
        label = CALL_LABELS[index] if index < len(CALL_LABELS) else 'extra call'
        gdb.write('\nOBSERVATION %d (%s)\n' % (index, label))
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

    @staticmethod
    def verdict(values):
        """Classify a per-call series against the fixed A, A, B, C call order.

        The repeat on shape A is what makes the classification possible: a value
        that changes between calls 0 and 1 cannot denote the Shape.
        """
        if len(values) != EXPECTED_OBSERVATIONS:
            return 'identical' if len(set(values)) == 1 else 'differs per call'
        same_shape = values[0] == values[1]
        same_slide = values[1] == values[2]
        all_equal = len(set(values)) == 1
        if all_equal:
            return 'constant across every call'
        if not same_shape:
            return 'differs per call (cannot denote the Shape)'
        if same_slide:
            return 'stable per slide'
        return 'stable per Shape'

    @staticmethod
    def verdict(values):
        """Classify a per-call series against the fixed A, A, B, C call order.

        The repeat on shape A is what makes the classification sound: a value
        that changes between calls 0 and 1 cannot denote the Shape, however
        neatly it happens to differ between different Shapes.
        """
        if len(values) != EXPECTED_OBSERVATIONS:
            return 'identical' if len(set(values)) == 1 else 'differs between calls'
        if len(set(values)) == 1:
            return 'constant across every call'
        if values[0] != values[1]:
            return 'differs per call, so it cannot denote the Shape'
        if values[1] == values[2]:
            return 'stable per slide'
        return 'stable per Shape'

    def summarize(self):
        """Classify every per-call value; say nothing beyond what varied."""
        if len(self.observations) < 2:
            gdb.write('\nToo few observations to compare.\n')
            return
        gdb.write('\nRECEIVER COMPARISON ACROSS %d CALLS\n' % len(self.observations))
        gdb.write('  call order: %s\n' % ', '.join(CALL_LABELS[:len(self.observations)]))
        for name in ['handler', 'token', 'receiver', 'cached', 'sequence']:
            values = [observation[name] for observation in self.observations]
            gdb.write('  %-16s %s\n      -> %s\n' % (
                name, ' '.join('%#x' % value for value in values),
                self.verdict(values)))
        for name in ['container', 'record_reference', 'reference_bytes']:
            values = [observation[name] for observation in self.observations]
            gdb.write('  %-16s -> %s\n' % (name, self.verdict(values)))
            for index, value in enumerate(values):
                gdb.write('      %-22s %s\n' % (CALL_LABELS[index] + ':', value))

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
                gdb.write('    +%#-6x %s\n           -> %s\n' % (
                    offset, ' '.join('%#x' % value for value in values),
                    self.verdict(values)))


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
        token = probe.pending_token or 0
        probe.record_observation({
            'handler': probe.pending_handler or 0,
            'token': token,
            'receiver': receiver,
            'cached': cached,
            'receiver_words': probe.words(receiver, RECEIVER_OBSERVATION_SIZE),
            'record_words': probe.words(transaction + PROPERTY_RECORD_OFFSET,
                                        RECORD_PREFIX_SIZE),
            'container': probe.described(receiver + RECEIVER_CONTAINER_OFFSET),
            'record_reference': probe.described(
                transaction + PROPERTY_RECORD_OFFSET + RECORD_REFERENCE_OFFSET),
            'sequence': probe.pointer(receiver + RECEIVER_SEQUENCE_OFFSET),
            'token_strong': probe.dword(token + TOKEN_STRONG_COUNT_OFFSET) if token else 0,
            'token_is_receiver': bool(
                token and probe.pointer(token + TOKEN_POINTEE_OFFSET) == receiver),
            'reference_bytes': probe.snapshot(
                probe.pointer(transaction + PROPERTY_RECORD_OFFSET
                              + RECORD_REFERENCE_OFFSET)),
        })
        return False


ReceiverProbe().execute()
