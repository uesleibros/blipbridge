# Run inside GDB's embedded Python; read-only hardware watchpoint on our image buffer.
import gdb, json, pathlib, struct
root = pathlib.Path.cwd()
target = json.loads((root / 'artifacts/decoder_target.json').read_text(encoding='utf-8-sig'))
gdb.execute('set pagination off')
gdb.execute('set confirm off')
gdb.execute('set solib-search-path C:/Program Files/Microsoft Office/root/vfs/ProgramFilesCommonX64/Microsoft Shared/Office16;C:/Program Files/Microsoft Office/root/Office16')
gdb.execute('attach ' + str(target['pid']))
expected = bytes.fromhex('45 33 c0 44 8b e0 41 8b c0')
if bytes(gdb.selected_inferior().read_memory(int(target['readReturn'],16),len(expected))) != expected:
    gdb.execute('detach')
    raise RuntimeError('MSO20 read-return signature mismatch')
class BufferFree(gdb.Breakpoint):
    def __init__(self, address):
        super().__init__('RtlFreeHeap', internal=True)
        self.address = address
    def stop(self):
        if int(gdb.parse_and_eval('$r8')) == self.address:
            gdb.write('\nWATCHED PNG BUFFER FREED; subsequent reuse is not image evidence\n')
            gdb.execute('bt 24')
            return True
        return False
class ImageReadWatch(gdb.Breakpoint):
    count = 0
    def stop(self):
        self.count += 1
        gdb.write('\nIMAGE BUFFER ACCESS %d\n' % self.count)
        gdb.execute('x/3i $pc')
        gdb.execute('bt 18')
        return self.count >= 12
def location(address):
    for m in target['modules']:
        if m['base'] <= address < m['base'] + m['size']:
            return '%s+0x%x' % (m['name'],address-m['base'])
    return hex(address)
def describe_object(address):
    inferior=gdb.selected_inferior()
    vtable=struct.unpack('<Q',bytes(inferior.read_memory(address,8)))[0]
    gdb.write('Object 0x%x vtable %s\n' % (address,location(vtable)))
    for i in range(12):
        f=struct.unpack('<Q',bytes(inferior.read_memory(vtable+i*8,8)))[0]
        gdb.write('  slot +0x%x = %s\n' % (i*8,location(f)))
    try:
        col=struct.unpack('<Q',bytes(inferior.read_memory(vtable-8,8)))[0]
        values=struct.unpack('<6I',bytes(inferior.read_memory(col,24)))
        if values[0]==1:
            typename=bytes(inferior.read_memory(col-values[5]+values[3]+16,256)).split(b'\0')[0]
            gdb.write('  RTTI '+repr(typename)+'\n')
    except gdb.MemoryError:
        gdb.write('  RTTI unavailable\n')
class SinkCall(gdb.Breakpoint):
    def stop(self):
        gdb.write('\nSOURCE-TO-SINK virtual call\n')
        for reg in ['rcx','rdx','r8','r9','r14']:
            gdb.write('%s=%s\n' % (reg,gdb.parse_and_eval('$'+reg)))
        describe_object(int(gdb.parse_and_eval('$rcx')))
        describe_object(int(gdb.parse_and_eval('$r14')))
        self.enabled=False
        return False
SinkCall('*'+target['sinkCall'],internal=True)
class ReadReturn(gdb.Breakpoint):
    def stop(self):
        address = int(gdb.parse_and_eval('$r13'))
        try:
            header = bytes(gdb.selected_inferior().read_memory(address, 8))
        except gdb.MemoryError:
            return False
        if header != b'\x89PNG\r\n\x1a\n':
            return False
        gdb.write('Watching PNG buffer at 0x%x\n' % address)
        self.enabled = False
        ImageReadWatch('*((unsigned long long*)0x%x)' % address,
                       gdb.BP_WATCHPOINT, wp_class=gdb.WP_ACCESS, internal=True)
        BufferFree(address)
        return False
ReadReturn('*' + target['readReturn'], internal=True)
(root / 'artifacts/decoder_go').write_text('go')
gdb.execute('continue')
for bp in gdb.breakpoints():
    bp.delete()
gdb.execute('detach')
