"""Run with GDB -batch -x after prepare_decoder_trace.ps1. No inferior calls.
For factory entry/return mode, precede -x with -ex "python probe_mode='factory'".
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
    factory=globals().get('probe_mode')=='factory'
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
