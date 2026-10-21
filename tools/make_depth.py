# 用 Depth Anything V2 (small, onnx) 为每张站点影像生成深度图
# 输出：assets/depth/<name>.jpg（灰度，越亮越近）
import io, os, sys, json
import numpy as np
from PIL import Image
import onnxruntime as ort

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODEL = os.path.join(ROOT, "tools", "depth_model.onnx")
SRC = os.path.join(ROOT, "assets", "img")
OUT = os.path.join(ROOT, "assets", "depth")
os.makedirs(OUT, exist_ok=True)

MEAN = np.array([0.485, 0.456, 0.406], dtype=np.float32)
STD = np.array([0.229, 0.224, 0.225], dtype=np.float32)
SIZE = 518
WIDTH = 1024

sess = ort.InferenceSession(MODEL, providers=["CPUExecutionProvider"])
in_name = sess.get_inputs()[0].name

def run(path):
    img = Image.open(path).convert("RGB")
    w, h = img.size
    # keep aspect, resize so the short side is 518
    if w >= h:
        nw, nh = SIZE, max(14, int(round(SIZE * h / w / 14)) * 14)
    else:
        nh, nw = SIZE, max(14, int(round(SIZE * w / h / 14)) * 14)
    im = img.resize((nw, nh), Image.BICUBIC)
    x = np.asarray(im, dtype=np.float32) / 255.0
    x = (x - MEAN) / STD
    x = np.transpose(x, (2, 0, 1))[None, ...].astype(np.float32)
    d = sess.run(None, {in_name: x})[0]
    d = np.squeeze(d)
    if d.ndim == 3:
        d = d[0]
    # 百分位归一化，避免极值把画面压平
    lo, hi = np.percentile(d, 1), np.percentile(d, 99)
    d = np.clip((d - lo) / max(1e-6, hi - lo), 0, 1)
    out_h = max(1, int(round(WIDTH * h / w)))
    dimg = Image.fromarray((d * 255).astype(np.uint8), mode="L").resize((WIDTH, out_h), Image.BICUBIC)
    name = os.path.splitext(os.path.basename(path))[0]
    dst = os.path.join(OUT, name + ".jpg")
    dimg.save(dst, quality=82, optimize=True)
    return dst, os.path.getsize(dst)

targets = sys.argv[1:]
if not targets:
    targets = [f for f in sorted(os.listdir(SRC)) if f.lower().endswith((".jpg", ".png"))]
for f in targets:
    p = os.path.join(SRC, f)
    if not os.path.exists(p):
        print("skip", f)
        continue
    try:
        dst, size = run(p)
        print("ok", os.path.basename(dst), size)
    except Exception as e:
        print("fail", f, repr(e))
