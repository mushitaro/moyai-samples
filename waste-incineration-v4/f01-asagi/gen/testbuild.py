import time, os, glob, sys
import ans
for f in glob.glob(ans.FILES + '/*'):
    os.remove(f)
t=time.time()
import build_sd as B, world2 as W
for fy in [int(a) for a in sys.argv[1:]]:
    B.sched_teishu(fy); B.daily_teishu(fy)
    for c, L in W.hoshu_all(fy):
        B.contract_docs(fy, c, L, 'H')
print('built', time.time()-t)
B.process_photos()
print('photos', time.time()-t)
fs=os.listdir(ans.FILES); print(len(fs)); print([f for f in sorted(fs) if f.endswith('xlsx')])
import statistics
js=[os.path.getsize(ans.FILES+'/'+f) for f in fs if f.endswith('jpg')]
if js: print(len(js), min(js), statistics.mean(js), max(js))
