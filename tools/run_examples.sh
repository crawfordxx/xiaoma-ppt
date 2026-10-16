#!/bin/bash
# 重跑全部示例：生成 → 渲染 → lint。
# 加 --replan：从头拆模板、出草稿（覆盖 deck.json），再用 refine_examples.py 重新精修。
set -e
cd "$(dirname "$0")/.."
EXAMPLES="finance:finance-ruixin medical:medical-kangning education:education-qiming manufacturing:manufacturing-huazhu tourism:tourism-yunqi"
if [ "$1" = "--replan" ]; then
  for pair in $EXAMPLES; do
    ex=${pair%%:*}; tp=${pair##*:}; d=examples/$ex
    python3 engine/inspect_template.py examples/templates/$tp.pptx --out $d/work >/dev/null
    python3 engine/plan.py $d/content.md $d/work/profile.json --out $d/deck.json >/dev/null
  done
  python3 tools/refine_examples.py >/dev/null
fi
for pair in $EXAMPLES; do
  ex=${pair%%:*}; d=examples/$ex
  python3 engine/build.py $d/deck.json | sed "s/^/[$ex] /" | grep -v 已生成 || true
  python3 engine/render.py $d/out/deck.pptx --out $d/out/render --dpi 110 >/dev/null
  python3 engine/lint.py $d/out/deck.pptx | tail -1 | sed "s/^/[$ex] lint /"
done
