"""Run with GDB -batch -x after prepare_decoder_trace.ps1. No inferior calls.
Precede -x with -ex "python probe_mode='factory'" for entry/return inspection.
Use 'consumer' to inspect OART retention, or 'lifetime' to follow references until
the first tracked final release. See docs/resource_lifetime.md for the protocol.
"""
import gdb, json, pathlib, struct

# Observational profile: GFX 16.0.14334.20848 x64 only. These addresses are
# breakpoint locations, never callable function pointers. Entry signatures were
# captured by objdump and verified dynamically. Evidence:
# docs/evidence/decoder_stream.txt, docs/evidence/cached_factory.txt,
# docs/userpicture_pipeline.md.
OFFICE_BUILD = '16.0.14334.20848'
CACHED_IMAGE_CREATE_RVA = 0x7680
CACHED_IMAGE_CREATE_SIGNATURE = bytes.fromhex('48 89 4c 24 08 53 48 83 ec 30')
STREAM_READ_RVA = 0x7880
STREAM_READ_SIGNATURE = bytes.fromhex('4c 8b dc 49 89 5b 08 49 89 6b 10')

# GFX stream adapter's borrowed IStream member, verified by the underlying
# MSO20 QueryInterface recognizing IID_IStream. This offset is not universal.
UNDERLYING_STREAM_OFFSET = 0x18
POINTER_BYTES = 8

# Windows x64 factory entry: return address + shadow space precede arguments
# five (MD4UID pointer) and six (bool). Includes hidden return storage in RCX.
FACTORY_UID_STACK_OFFSET = 0x28
FACTORY_BOOL_STACK_OFFSET = 0x30

# OART 16.0.14334.20848 x64 consumer observed directly after factory return.
# The +0xF0 member is passed to a reference assignment helper; it is not yet
# proven to be a document BLIP. The +8 count belongs to private GFX objects,
# whose slot 0 increments and slot +8 decrements (NOT IUnknown layout).
OART_CONSUMER_RVA = 0x8F94C
OART_CONSUMER_SIGNATURE = bytes.fromhex('48 89 5c 24 20 56 57 41 56')
OART_AFTER_LOCAL_RELEASE_RVA = 0x8F575
OART_AFTER_LOCAL_RELEASE_SIGNATURE = bytes.fromhex('48 8d 4d df e8')
OART_CACHE_MEMBER_OFFSET = 0xF0
GFX_REFERENCE_COUNT_OFFSET = 8
GFX_CACHED_RELEASE_RVA = 0xA1A0
GFX_IMAGE_RELEASE_RVA = 0xA170
GFX_RELEASE_SIGNATURE = bytes.fromhex('83 c8 ff f0 0f c1 41 08')
# Shared intrusive increment entry, observed on both returned GFX interfaces.
GFX_ADD_REFERENCE_RVA = 0x90AA0
GFX_ADD_REFERENCE_SIGNATURE = bytes.fromhex('f0 ff 41 08 c3')

root = pathlib.Path.cwd()
target = json.loads((root/'artifacts/decoder_target.json').read_text(encoding='utf-8-sig'))
gdb.execute('set pagination off')
gdb.execute('set confirm off')
gdb.execute('set solib-search-path C:/Program Files/Microsoft Office/root/vfs/ProgramFilesCommonX64/Microsoft Shared/Office16;C:/Program Files/Microsoft Office/root/Office16')
gdb.execute('attach '+str(target['pid']))
def read(address, size):
    return bytes(gdb.selected_inferior().read_memory(address, size))
def ptr(address):
    """Read a borrowed x64 address; never retain or dereference it after detach."""
    return struct.unpack('<Q',read(address,POINTER_BYTES))[0]
def location(address):
    for module in target['modules']:
        if module['base'] <= address < module['base']+module['size']:
            return '%s+0x%x'%(module['name'],address-module['base'])
    return hex(address)
def describe(address):
    """Report vtable entries as observations, without assuming IUnknown layout."""
    vtable=ptr(address)
    gdb.write('object=%#x vtable=%s\n'%(address,location(vtable)))
    for slot in range(14):
        gdb.write('  +%#x %s\n'%(slot*8,location(ptr(vtable+slot*8))))
class StreamRead(gdb.Breakpoint):
    def stop(self):
        adapter=int(gdb.parse_and_eval('$rcx'))
        gdb.write('\nGFX decoder adapter\n')
        describe(adapter)
        gdb.write('Underlying stream at adapter+0x18\n')
        describe(ptr(adapter+UNDERLYING_STREAM_OFFSET))
        gdb.execute('bt 18')
        return True
class FactoryReturn(gdb.Breakpoint):
    def __init__(self, address, storage, image_out, thread):
        super().__init__('*'+hex(address),internal=True)
        self.storage, self.image_out, self.thread_id = storage, image_out, thread
    def stop(self):
        if gdb.selected_thread().ptid != self.thread_id:
            return False
        gdb.write('\nFACTORY RETURN rax=%s\n'%gdb.parse_and_eval('$rax'))
        for label, storage in [('cached result',self.storage),('image out',self.image_out)]:
            value=ptr(storage)
            gdb.write('%s storage=%#x value=%#x\n'%(label,storage,value))
            if value:
                describe(value)
        gdb.execute('bt 12')
        if globals().get('probe_mode') in ['consumer', 'lifetime']:
            self.enabled = False
            observation = ConsumerObservation(ptr(self.storage), ptr(self.image_out))
            observation.install()
            return False
        return True


class ConsumerObservation:
    """Borrows addresses only while stopped; never calls AddRef/Release itself."""

    def __init__(self, cached_image, image):
        self.cached_image = cached_image
        self.image = image
        self.owner = None
        self.thread_id = gdb.selected_thread().ptid

    def report(self, stage):
        gdb.write('\nCONSUMER ' + stage + '\n')
        for label, address in [('cached', self.cached_image), ('image', self.image)]:
            count = struct.unpack('<I', read(address + GFX_REFERENCE_COUNT_OFFSET, 4))[0]
            gdb.write('%s object=%#x count=%d\n' % (label, address, count))
        if self.owner is not None:
            member = self.owner + OART_CACHE_MEMBER_OFFSET
            gdb.write('OART record=%#x first-qword=%s member+0xF0=%s\n' % (
                self.owner, location(ptr(self.owner)), read(member, 32).hex()))

    def install(self):
        gfx = next(module for module in target['modules'] if module['name'].lower() == 'gfx.dll')
        oart = next(module for module in target['modules'] if module['name'].lower() == 'oart.dll')
        if oart.get('version') != OFFICE_BUILD:
            raise RuntimeError('Unsupported OART build')
        profiles = [
            (oart, OART_CONSUMER_RVA, OART_CONSUMER_SIGNATURE),
            (oart, OART_AFTER_LOCAL_RELEASE_RVA, OART_AFTER_LOCAL_RELEASE_SIGNATURE),
            (gfx, GFX_CACHED_RELEASE_RVA, GFX_RELEASE_SIGNATURE),
            (gfx, GFX_IMAGE_RELEASE_RVA, GFX_RELEASE_SIGNATURE),
            (gfx, GFX_ADD_REFERENCE_RVA, GFX_ADD_REFERENCE_SIGNATURE),
        ]
        for module, rva, signature in profiles:
            if read(module['base'] + rva, len(signature)) != signature:
                raise RuntimeError('Consumer/release signature mismatch')
        ConsumerEntry(oart['base'] + OART_CONSUMER_RVA, self)
        ConsumerComplete(oart['base'] + OART_AFTER_LOCAL_RELEASE_RVA, self)
        for rva in [GFX_CACHED_RELEASE_RVA, GFX_IMAGE_RELEASE_RVA]:
            ObservedRelease(gfx['base'] + rva, self)
        ObservedAddReference(gfx['base'] + GFX_ADD_REFERENCE_RVA, self)
        self.report('factory return')


class ConsumerEntry(gdb.Breakpoint):
    def __init__(self, address, observation):
        super().__init__('*' + hex(address), internal=True)
        self.observation = observation

    def stop(self):
        if gdb.selected_thread().ptid != self.observation.thread_id:
            return False
        self.observation.owner = int(gdb.parse_and_eval('$rcx'))
        self.observation.report('before OART consumer')
        self.enabled = False
        return False


class ObservedRelease(gdb.Breakpoint):
    def __init__(self, address, observation):
        super().__init__('*' + hex(address), internal=True)
        self.observation = observation

    def stop(self):
        address = int(gdb.parse_and_eval('$rcx'))
        if address in [self.observation.cached_image, self.observation.image]:
            count = struct.unpack('<I', read(address + GFX_REFERENCE_COUNT_OFFSET, 4))[0]
            gdb.write('RELEASE entry object=%#x count-before=%d caller=%s\n' % (
                address, count, location(ptr(int(gdb.parse_and_eval('$rsp'))))))
            phase_file = root / 'artifacts/decoder_phase.txt'
            if phase_file.exists():
                gdb.write('Harness phase: ' + phase_file.read_text().strip() + '\n')
            # Never continue to read a tracked object after its last release.
            if count == 1:
                gdb.write('Last release observed; stopping before potential destruction\n')
                gdb.execute('bt 24')
                return True
        return False


class ObservedAddReference(gdb.Breakpoint):
    """Capture a bounded set of cached-resource retention callers, read-only."""

    def __init__(self, address, observation):
        super().__init__('*' + hex(address), internal=True)
        self.observation = observation
        self.remaining_stacks = 12

    def stop(self):
        address = int(gdb.parse_and_eval('$rcx'))
        if address != self.observation.cached_image or self.remaining_stacks == 0:
            return False
        self.remaining_stacks -= 1
        count = struct.unpack('<I', read(address + GFX_REFERENCE_COUNT_OFFSET, 4))[0]
        gdb.write('\nCACHED ADDREF count-before=%d\n' % count)
        gdb.execute('bt 10')
        return False


class ConsumerComplete(gdb.Breakpoint):
    def __init__(self, address, observation):
        super().__init__('*' + hex(address), internal=True)
        self.observation = observation

    def stop(self):
        if gdb.selected_thread().ptid != self.observation.thread_id:
            return False
        self.observation.report('after OART consumer and local cached release')
        if globals().get('probe_mode') == 'lifetime':
            self.enabled = False
            return False
        return True
class FactoryEntry(gdb.Breakpoint):
    def stop(self):
        regs={r:int(gdb.parse_and_eval('$'+r)) for r in ['rcx','rdx','r8','r9','rsp']}
        gdb.write('\nFACTORY ENTRY '+str(regs)+'\n')
        gdb.write('optional UID pointer=%#x bool=%d\n'%(
            ptr(regs['rsp']+FACTORY_UID_STACK_OFFSET),
            read(regs['rsp']+FACTORY_BOOL_STACK_OFFSET,1)[0]))
        gdb.write('Input stream\n')
        describe(regs['r8'])
        FactoryReturn(ptr(regs['rsp']),regs['rcx'],regs['rdx'],gdb.selected_thread().ptid)
        self.enabled=False
        return False
try:
    gfx=next(m for m in target['modules'] if m['name'].lower()=='gfx.dll')
    if gfx.get('version') != OFFICE_BUILD:
        raise RuntimeError('Unsupported GFX build or missing module version')
    factory=globals().get('probe_mode') in ['factory', 'consumer', 'lifetime']
    address=gfx['base']+(CACHED_IMAGE_CREATE_RVA if factory else STREAM_READ_RVA)
    signature=CACHED_IMAGE_CREATE_SIGNATURE if factory else STREAM_READ_SIGNATURE
    if read(address,len(signature))!=signature:
        raise RuntimeError('GFX entry signature mismatch')
    (FactoryEntry if factory else StreamRead)('*'+hex(address),internal=True)
    (root/'artifacts/decoder_go').write_text('go')
    gdb.execute('continue')
finally:
    for breakpoint in gdb.breakpoints() or []:
        breakpoint.delete()
    gdb.execute('detach')
