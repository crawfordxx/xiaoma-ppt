"""回归测试：python3 -m pytest tests/ -q（不需要 LibreOffice，渲染相关的用 --no-render）"""
import json
import os
import subprocess
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENGINE = os.path.join(ROOT, "engine")
sys.path.insert(0, ENGINE)

from plan import outline, parse, unit_slides  # noqa: E402

TEMPLATES = {"finance-ruixin": [0, 1, 2, 3, 4, 5], "medical-kangning": [0, 1, 2, 3, 4],
             "education-qiming": [0, 1, 2, 3, 4], "manufacturing-huazhu": [0, 1, 2, 3, 4, 5],
             "tourism-yunqi": [0, 1, 2, 3, 4]}
EXPECT_ROLES = {"finance-ruixin": ["cover", "toc", "section", "content", "content", "ending"],
                "medical-kangning": ["cover", "toc", "section", "content", "ending"],
                "education-qiming": ["cover", "toc", "section", "content", "ending"],
                "manufacturing-huazhu": ["cover", "toc", "section", "content", "content", "ending"],
                "tourism-yunqi": ["cover", "toc", "section", "content", "ending"]}


def run(*args, cwd=ROOT):
    return subprocess.run([sys.executable, *args], cwd=cwd, capture_output=True, text=True)


@pytest.fixture(scope="module")
def profiles(tmp_path_factory):
    out = {}
    for name in TEMPLATES:
        d = tmp_path_factory.mktemp(name)
        r = run(os.path.join(ENGINE, "inspect_template.py"), f"examples/templates/{name}.pptx", "--out", str(d),
                "--no-render")
        assert r.returncode == 0, r.stderr
        out[name] = (d, json.load(open(d / "profile.json", encoding="utf-8")))
    return out


@pytest.mark.parametrize("name", list(TEMPLATES))
def test_roles(profiles, name):
    _, prof = profiles[name]
    assert [s["role"] for s in prof["slides"]] == EXPECT_ROLES[name]


def test_brand_text_is_not_a_slot(profiles):
    _, prof = profiles["finance-ruixin"]
    cover = prof["slides"][0]
    brand = [t["key"] for t in cover["texts"] if t["text"] == "睿信资本"]
    assert brand and brand[0] not in sum([v if isinstance(v, list) else [v] for v in cover["slots"].values()], [])


def test_toc_items_found(profiles):
    for name in TEMPLATES:
        toc = profiles[name][1]["slides"][1]
        assert len(toc["slots"]["items"]) >= 4, name


def blocks(md):
    return outline(parse(md), "fallback")


def test_bom_title(profiles, tmp_path):
    md = tmp_path / "bom.md"
    md.write_bytes("# 带 BOM 的标题\n\n## 一\n\n- a\n- b\n".encode("utf-8-sig"))
    d, _ = profiles["finance-ruixin"]
    r = run(os.path.join(ENGINE, "plan.py"), str(md), str(d / "profile.json"), "--out", str(tmp_path / "deck.json"))
    assert r.returncode == 0, r.stderr
    assert json.load(open(tmp_path / "deck.json", encoding="utf-8"))["slides"][0]["title"] == "带 BOM 的标题"


def test_no_h2_keeps_content():
    doc = blocks("# T\n\n### 页一\n\n- **a**：1\n- **b**：2\n")
    assert doc["sections"] and doc["sections"][0]["units"], "没有 ## 的内容不能被丢掉"


def test_list_does_not_swallow_heading():
    b = parse("- a\n- b\n### Q\ntext\n")
    assert b[0]["items"] == ["a", "b"] and b[1]["t"] == "h"


def test_loose_list_is_one_block():
    b = parse("- a\n\n- b\n\n- c\n")
    assert len(b) == 1 and b[0]["items"] == ["a", "b", "c"]


def test_image_with_title_and_url():
    b = parse('![x](a.png "标题")\n\n![y](https://e.com/b.png)\n')
    assert [x.get("path") for x in b] == ["a.png"]


def test_long_table_paged():
    rows = "\n".join(f"| 项{i} | 说明{i} |" for i in range(20))
    doc = blocks(f"# T\n\n## S\n\n| 名称 | 说明 |\n|---|---|\n{rows}\n")
    out = unit_slides(doc["sections"][0]["units"][0])
    assert len(out) == 3 and all(o[0]["type"] == "table" and len(o[0]["rows"]) <= 8 for o in out)


def test_compare_followed_by_table_keeps_table():
    md = "# T\n\n## S\n\n**过去**\n\n- a\n- b\n\n**现在**\n\n- c\n- d\n\n| 年 | 值 |\n|---|---|\n| 2024 | 1 |\n| 2025 | 2 |\n"
    out = unit_slides(blocks(md)["sections"][0]["units"][0])
    assert [o[0]["type"] for o in out] == ["compare", "chart"]


def test_stats_uses_bold_number():
    out = unit_slides(blocks("# T\n\n## S\n\n- 国家 4A 级以上景区 **9 处**\n- 森林覆盖率 **78%**\n")["sections"][0]["units"][0])
    items = out[0][0]["items"]
    assert out[0][0]["type"] == "stats" and items[0]["value"] == "9" and items[0]["unit"] == "处"


@pytest.mark.parametrize("ex", ["finance", "medical", "education", "manufacturing", "tourism"])
def test_examples_build_and_lint(ex, tmp_path):
    deck = os.path.join(ROOT, "examples", ex, "deck.json")
    out = tmp_path / "deck.pptx"
    r = run(os.path.join(ENGINE, "build.py"), deck, "--out", str(out), cwd=str(tmp_path))  # 在别的目录运行
    assert r.returncode == 0, r.stderr
    r = run(os.path.join(ENGINE, "lint.py"), str(out), "--no-render")
    assert r.returncode == 0, r.stdout
