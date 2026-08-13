#!/usr/bin/env python3
"""Word -> Twee -> standalone SugarCube HTML converter.

Supports three compatible authoring modes in the same DOCX:
1) Simple branching: H1 story title, H2 passages, [[label->Passage]] links.
2) Recommended simplified directives: Choice/MultiChoice/WhenChosen/etc.
3) Legacy advanced directives: id/type/label/target syntax remains supported.

The HTML builder contains an embedded SugarCube runtime shell, so Twine, Tweego,
or a separate template HTML file are not required at conversion time.
"""
from __future__ import annotations

import argparse
import html
import json
import math
import re
import sys
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document

DIRECTIVE_RE = re.compile(r"^\s*([A-Za-z][A-Za-z0-9_-]*)\s*::\s*(.*?)\s*$")
LINK_RE = re.compile(r"\[\[(.*?)(?:->|\|)(.*?)\]\]")

@dataclass
class Para:
    text: str
    style: str

@dataclass
class Passage:
    title: str
    items: List[Para] = field(default_factory=list)

@dataclass
class Story:
    title: str
    start: str
    theme: str = "literature-review-light"
    passages: List[Passage] = field(default_factory=list)
    config: Dict[str, str] = field(default_factory=dict)


def parse_kv(payload: str) -> Tuple[Optional[str], Dict[str, str]]:
    parts = [p.strip() for p in payload.split("|")]
    head = None
    out: Dict[str, str] = {}
    for i, part in enumerate(parts):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.strip()
        elif i == 0 and part:
            head = part
    return head, out


def style_name(p) -> str:
    try:
        return p.style.name or "Normal"
    except Exception:
        return "Normal"


STYLE_DIRECTIVE_MAP = {
    "Scenario Choice": "Choice",
    "Scenario MultiChoice": "MultiChoice",
    "Scenario Cost": "Cost",
    "Scenario Set": "Set",
    "Scenario Add": "Add",
    "Scenario Feedback": "Feedback",
}

def para_from_docx(p) -> Optional[Para]:
    """Convert a Word paragraph into the converter's internal Para model.

    The recommended authoring format uses Word styles to carry common
    semantics, so authors do not need to type Choice ::, Set ::, Cost ::,
    etc. Legacy directive-based documents remain supported unchanged.
    """
    raw = p.text.strip()
    if not raw:
        return None

    style = style_name(p)
    directive = STYLE_DIRECTIVE_MAP.get(style)

    if directive:
        # If an older document already contains the directive text, do not
        # duplicate it.
        if DIRECTIVE_RE.match(raw):
            return Para(raw, style)
        return Para(f"{directive} :: {raw}", style)

    return Para(raw, style)


def load_docx(path: Path) -> Story:
    doc = Document(path)
    paras = []
    for p in doc.paragraphs:
        converted = para_from_docx(p)
        if converted is not None:
            paras.append(converted)

    config: Dict[str, str] = {}
    declared_title = None
    start = None
    theme = "literature-review-light"
    for p in paras:
        m = DIRECTIVE_RE.match(p.text)
        if not m:
            continue
        key, value = m.group(1), m.group(2)
        lk = key.lower()
        if lk == "scenariotitle":
            declared_title = value
        elif lk == "startpassage":
            start = value
        elif lk == "theme":
            theme = value
        elif lk in {
            "scenariosubtitle", "storyformat",
            "budgetvariable", "budget",
            "statusvariable", "status",
            "scoredimension", "score",
            "saveprogress", "restartenabled",
            "showprogressbar", "showsidebar", "accessibilitymode"
        }:
            config[key] = value

    h1s = [p.text for p in paras if p.style.startswith("Heading 1")]
    title = declared_title or (h1s[0] if h1s else path.stem)

    # Advanced prototype may contain guidance before a second H1 matching ScenarioTitle.
    start_idx = 0
    matching_h1 = [i for i,p in enumerate(paras) if p.style.startswith("Heading 1") and p.text == title]
    if matching_h1:
        start_idx = matching_h1[-1] + 1
    elif h1s:
        start_idx = next(i for i,p in enumerate(paras) if p.style.startswith("Heading 1")) + 1

    passages: List[Passage] = []
    current: Optional[Passage] = None
    for p in paras[start_idx:]:
        if p.style.startswith("Heading 1"):
            # Any later H1 is treated as author/development material, not story content.
            break
        if p.style.startswith("Heading 2"):
            current = Passage(p.text)
            passages.append(current)
        elif current is not None:
            current.items.append(p)

    if not passages:
        # Fallback: permit a basic document where H1 is title and H2 passages.
        current = None
        for p in paras:
            if p.style.startswith("Heading 2"):
                current = Passage(p.text); passages.append(current)
            elif current is not None and not p.style.startswith("Heading 1"):
                current.items.append(p)

    if not passages:
        raise ValueError("No Heading 2 passages found in the Word document.")
    start = start or passages[0].title
    return Story(title=title, start=start, theme=theme, passages=passages, config=config)


def sugar_expr(expr: str) -> str:
    """Convert readable scenario expressions to SugarCube syntax.

    Authors may use AND / OR / NOT; legacy && / || / ! also remain valid.
    """
    expr = expr.strip()
    expr = re.sub(r"\bAND\b", "&&", expr, flags=re.I)
    expr = re.sub(r"\bOR\b", "||", expr, flags=re.I)
    expr = re.sub(r"\bNOT\b\s+", "!", expr, flags=re.I)
    # Protect quoted strings.
    chunks = re.split(r"('(?:[^'\\]|\\.)*'|\"(?:[^\"\\]|\\.)*\")", expr)
    keywords = {"true","false","null","undefined","if","else","and","or","not","Math","ceil","max","min"}
    out = []
    for i, chunk in enumerate(chunks):
        if i % 2:
            out.append(chunk); continue
        def repl(m):
            w = m.group(0)
            if w in keywords or w.startswith("$") or re.fullmatch(r"\d+(?:\.\d+)?", w):
                return w
            return "$" + w
        chunk = re.sub(r"(?<![$.])\b[A-Za-z_][A-Za-z0-9_]*\b", repl, chunk)
        chunk = chunk.replace("$true","true").replace("$false","false").replace("$null","null")
        out.append(chunk)
    return "".join(out)


def setter_expr(expr: str) -> str:
    return sugar_expr(expr.replace(";", "; "))


def render_plain_para(item: Para) -> str:
    text = item.text.strip()
    if not text:
        return ""

    # Word Heading 3 should become a real HTML heading.  Also tolerate
    # literal Markdown-style ### headings in older authoring documents.
    if item.style.startswith("Heading 3"):
        return f"<h3>{html.escape(text)}</h3>"
    if text.startswith("### "):
        return f"<h3>{html.escape(text[4:].strip())}</h3>"

    if item.style == "Author Note":
        return f"<div class=\"author-note\">{html.escape(text)}</div>"

    # List paragraphs are grouped into <ul>/<ol> by render_passage().
    if item.style.startswith("List Bullet") or item.style.startswith("List Number"):
        return html.escape(text)

    # Ordinary Word paragraphs become explicit HTML paragraphs.  SugarCube
    # link syntax such as [[Continue->Next]] remains intact because square
    # brackets do not require HTML escaping.
    return f"<p>{html.escape(text)}</p>"


def choice_to_markup(payload: str, forced_type: Optional[str] = None) -> Tuple[str, str, str, str, str]:
    """Parse both simplified and legacy choice syntax.

    Recommended:
        Choice :: Review the evidence -> Evidence
        MultiChoice :: Check records -> Results | cost=4

    Legacy:
        Choice :: id=x | type=multi | label=Check records | target=Results | cost=4
    """
    head, kv = parse_kv(payload)

    # Simplified arrow syntax lives in the first/head segment.
    label = None
    target = None
    if head and "->" in head:
        label, target = [x.strip() for x in head.split("->", 1)]

    label = kv.get("label", label or head or "Continue")
    target = kv.get("target", target or "")
    typ = forced_type or kv.get("type", "single")
    cid = kv.get("id", re.sub(r"\W+", "_", label).strip("_").lower())
    cost = kv.get("cost", "")
    return cid, label, target, typ, cost


def render_passage(p: Passage) -> str:
    out: List[str] = []
    i = 0
    pending_choice: Optional[Tuple[str,str,str,str,str]] = None
    multi_choices: List[dict] = []
    multi_target = ""
    multi_cost_rule = "selected-total"
    gate = ""
    gate_fail = ""
    table_open = False
    outcome_open = False

    def close_table():
        nonlocal table_open
        if table_open:
            out.append("</tbody></table>")
            table_open = False

    def flush_multi():
        nonlocal multi_choices, multi_target
        if not multi_choices:
            return
        spec = html.escape(json.dumps(multi_choices), quote=True)
        target = html.escape(multi_target or p.title, quote=True)
        cr = html.escape(multi_cost_rule, quote=True)
        g = html.escape(gate, quote=True)
        gf = html.escape(gate_fail, quote=True)
        out.append(f'<button class="scenario-continue" onclick="ScenarioEngine.submitMulti(this, \'{target}\', JSON.parse(this.dataset.spec), \'{cr}\', \'{g}\', \'{gf}\')" data-spec="{spec}">Continue</button>')
        multi_choices = []
        multi_target = ""

    while i < len(p.items):
        item = p.items[i]

        # Allow structural conditional markers without '::' in Word.
        # Example:
        #   If :: condition
        #   ...
        #   Else
        #   ...
        #   EndIf
        bare = item.text.strip().lower()
        if bare == "else":
            close_table()
            out.append("<<else>>")
            i += 1
            continue
        if bare == "endif":
            close_table()
            out.append("<</if>>")
            i += 1
            continue

        # Convert consecutive Word bullet/number paragraphs into a single
        # semantic HTML list instead of emitting Markdown-like lines.
        if item.style.startswith("List Bullet"):
            close_table()
            items = []
            while i < len(p.items) and p.items[i].style.startswith("List Bullet") and not DIRECTIVE_RE.match(p.items[i].text):
                items.append(f"<li>{html.escape(p.items[i].text.strip())}</li>")
                i += 1
            out.append("<ul class=\"scenario-list\">" + "".join(items) + "</ul>")
            continue

        if item.style.startswith("List Number"):
            close_table()
            items = []
            while i < len(p.items) and p.items[i].style.startswith("List Number") and not DIRECTIVE_RE.match(p.items[i].text):
                items.append(f"<li>{html.escape(p.items[i].text.strip())}</li>")
                i += 1
            out.append("<ol class=\"scenario-list\">" + "".join(items) + "</ol>")
            continue

        m = DIRECTIVE_RE.match(item.text)
        if not m:
            close_table()
            rendered = render_plain_para(item)
            if rendered:
                out.append(rendered)
            i += 1
            continue
        key, payload = m.group(1).lower(), m.group(2)

        if key == "node":
            pass
        elif key == "set":
            assigns = []
            for part in payload.split(";"):
                part = part.strip()
                if not part: continue
                if re.match(r"^[A-Za-z_]\w*\s*=", part):
                    var, rhs = part.split("=",1)
                    assigns.append(f"${var.strip()} = {sugar_expr(rhs.strip())}")
                else:
                    assigns.append(sugar_expr(part))
            out.append(f"<<set {'; '.join(assigns)}>>")
        elif key == "add":
            out.append(f"<<set {sugar_expr(payload)}>>")
        elif key == "if":
            close_table(); out.append(f"<<if {sugar_expr(payload)}>>")
        elif key == "elseif":
            out.append(f"<<elseif {sugar_expr(payload)}>>")
        elif key == "else":
            out.append("<<else>>")
        elif key == "endif":
            out.append("<</if>>")
        elif key in {"choice", "multichoice"}:
            close_table()
            forced_type = "multi" if key == "multichoice" else None
            cid,label,target,typ,cost = choice_to_markup(payload, forced_type=forced_type)
            onchoose_parts = []

            # Recommended readable block:
            #   MultiChoice :: Review evidence -> Results
            #   Cost :: 4
            #   Set :: reviewed=true
            #   Add :: score += 2
            #
            # Legacy WhenChosen/OnChoose is still accepted.
            j = i + 1
            while j < len(p.items):
                nxt = p.items[j]

                if nxt.style == "Author Note":
                    j += 1
                    continue

                nm = DIRECTIVE_RE.match(nxt.text)
                if not nm:
                    break

                nk = nm.group(1).lower()
                nv = nm.group(2).strip()

                if nk in {"onchoose", "whenchosen"}:
                    onchoose_parts.append(nv)
                    j += 1
                    continue

                if nk == "cost":
                    cost = nv
                    j += 1
                    continue

                if nk == "set":
                    onchoose_parts.append(nv)
                    j += 1
                    continue

                if nk == "add":
                    onchoose_parts.append(nv)
                    j += 1
                    continue

                break

            if j > i + 1:
                i = j - 1

            onchoose = "; ".join(x for x in onchoose_parts if x)

            if typ == "multi":
                multi_target = multi_target or target
                multi_choices.append({
                    "id": cid,
                    "label": label,
                    "cost": float(cost) if cost else 0,
                    "onchoose": onchoose
                })
                out.append(
                    f'<label class="scenario-option">'
                    f'<input type="checkbox" class="scenario-multi" '
                    f'data-choice="{html.escape(cid)}">'
                    f'<span>{html.escape(label)}</span></label>'
                )
            else:
                setpart = f"[{setter_expr(onchoose)}]" if onchoose else ""
                out.append(f'<div class="scenario-action">[[{label}->{target}]{setpart}]</div>')
        elif key == "costrule":
            multi_cost_rule = payload.strip()
        elif key == "require":
            # Generic requirement text retained for validation/UI; multi forms require at least one choice automatically.
            out.append(f"<!-- Require :: {html.escape(payload)} -->")
        elif key == "gate":
            gate = payload.strip()
        elif key == "gatefailuretarget":
            gate_fail = payload.strip()
        elif key == "feedback":
            head, kv = parse_kv(payload)
            if kv.get("type") or kv.get("text"):
                typ = kv.get("type", "info")
                txt = kv.get("text", head or payload)
            else:
                # Simplified: Feedback :: good | Message text
                parts = [x.strip() for x in payload.split("|", 1)]
                typ = parts[0] if parts and parts[0] else "info"
                txt = parts[1] if len(parts) > 1 else payload
            out.append(f'<<script>>ScenarioEngine.feedback("{typ}", {json.dumps(txt)});<</script>>')
        elif key == "table":
            flush_multi(); close_table()
            _, kv = parse_kv(payload)
            headers = kv.get("headers", kv.get("columns", "")).split("|")
            out.append('<table class="scenario-table"><thead><tr>' + ''.join(f'<th>{html.escape(h)}</th>' for h in headers) + '</tr></thead><tbody>')
            table_open = True
        elif key == "row":
            if not table_open:
                out.append('<table class="scenario-table"><tbody>'); table_open=True
            cells = payload.split("|")
            out.append('<tr>' + ''.join(f'<td>{html.escape(c.strip())}</td>' for c in cells) + '</tr>')
        elif key == "chart":
            _, kv = parse_kv(payload)
            vals = [v.strip() for v in kv.get("data","").split(",") if v.strip()]
            if vals:
                try: nums=[float(v) for v in vals]; mx=max(nums) or 1
                except: nums=[]; mx=1
                bars=''.join(f'<div class="scenario-bar" style="height:{max(4, int(n/mx*140))}px" title="{n:g}"></div>' for n in nums)
                out.append(f'<div class="scenario-chart" role="img" aria-label="{html.escape(kv.get("y","Values"))} by {html.escape(kv.get("x","category"))}">{bars}</div>')
        elif key == "outcome":
            close_table(); flush_multi()
            head, kv = parse_kv(payload)

            # Recommended:
            #   Outcome :: Successful response | when=fixed && treated
            #   Outcome :: Uncontrolled | default
            # Legacy:
            #   Outcome :: if=fixed && treated | title=Successful response
            default_flag = any(part.strip().lower() == "default" for part in payload.split("|"))
            if "when" in kv:
                title = head or kv.get("title", "Outcome")
                cond = kv["when"]
            else:
                cond = kv.get("if", head or "")
                title = kv.get("title", "Outcome")

            if default_flag or cond.strip().lower() == "default":
                if outcome_open:
                    out.append("<<else>>")
                else:
                    # A default-only outcome needs no conditional wrapper.
                    pass
            else:
                if not outcome_open:
                    out.append(f"<<if {sugar_expr(cond)}>>")
                    outcome_open = True
                else:
                    out.append(f"<<elseif {sugar_expr(cond)}>>")
            out.append(f"<h3>{html.escape(title)}</h3>")
        elif key == "scoredisplay":
            head, kv = parse_kv(payload)
            label = head or "Score"
            value = kv.get("value", "score")
            maxv = kv.get("max", "")
            suffix = f" / {html.escape(maxv)}" if maxv else ""
            out.append(f'<div class="score-row"><strong>{html.escape(label)}:</strong> <<print ${value}>>{suffix}</div>')
        elif key == "restart":
            out.append('[[Restart->' + html.escape(p.title) + ']]')
        elif key == "calculate":
            # Supports simple assignment calculations using common
            # author-friendly functions mapped to JavaScript Math methods.
            if "=" in payload:
                var, rhs = payload.split("=", 1)
                rhs = sugar_expr(rhs.strip())
                rhs = rhs.replace("ceil(", "Math.ceil(")
                rhs = rhs.replace("floor(", "Math.floor(")
                rhs = rhs.replace("round(", "Math.round(")
                rhs = rhs.replace("max(", "Math.max(")
                rhs = rhs.replace("min(", "Math.min(")
                rhs = rhs.replace("abs(", "Math.abs(")
                out.append(f"<<set ${var.strip()} = {rhs}>>")
        elif key == "clamp":
            mm=re.match(r"(\w+)\s+between\s+([\d.]+)\s+and\s+([\d.]+)",payload,re.I)
            if mm:
                v,lo,hi=mm.groups(); out.append(f"<<set ${v} = Math.max({lo}, Math.min({hi}, ${v}))>>")
        elif key == "grade":
            out.append(f"<!-- Grade rule :: {html.escape(payload)} -->")
        elif key == "gradecap":
            out.append(f"<!-- GradeCap :: {html.escape(payload)} -->")
        elif key in {"criticalaction","debrief","replacecell","scorevariable","scoredimension"}:
            out.append(f"<!-- {m.group(1)} :: {html.escape(payload)} -->")
        else:
            out.append(f"<!-- Unsupported directive retained: {html.escape(item.text)} -->")
        i += 1

    close_table(); flush_multi()
    if outcome_open:
        out.append("<</if>>")

    # Compact passage markup. SugarCube 1.x can preserve repeated blank lines
    # between block elements as visible vertical whitespace.
    # Emit passage markup without source-level whitespace between blocks.
    # SugarCube 1.x can wikify otherwise harmless newlines around invisible
    # macros (Set/Add/Feedback/etc.) into visible vertical gaps. HTML block
    # elements and SugarCube macros do not require separating newlines.
    return "".join(x.strip() for x in out if x is not None and x.strip())


def story_to_twee(story: Story) -> str:
    lines = [f":: StoryTitle\n{story.title}\n"]
    # StoryInit is understood by SugarCube and is a safe place for feedback initialization.
    lines.append(":: StoryInit [script]\n<<set $scenarioFeedback = []>>\n")
    for p in story.passages:
        lines.append(f":: {p.title}\n{render_passage(p)}\n")
    return "\n".join(lines)


def extract_custom_css(template: str) -> str:
    m = re.search(r'<style role="stylesheet" id="twine-user-stylesheet" type="text/twine-css">(.*?)</style>', template, re.S)
    if m:
        return html.unescape(m.group(1))
    return """
html,body{background:#fff;color:#222;font-family:'Segoe UI',Arial,sans-serif;font-size:20px;line-height:1.7}
#ui-bar{background:#f5f5f5;color:#222;border-right:1px solid #d9d9d9}.passage{max-width:900px;margin:0 auto;padding:2rem}
a,.link-internal{color:#005a9c}.scenario-option{display:block;margin:.75rem 0;padding:.7rem;border:1px solid #d9d9d9;border-radius:5px}
"""

SCENARIO_JS = r'''
window.ScenarioEngine = (function () {
  function vars(){ return state.active.variables; }
  function value(s){
    s=(s||'').trim();
    if(/^[-+]?\d+(\.\d+)?$/.test(s)) return Number(s);
    if(s==='true') return true; if(s==='false') return false; if(s==='null') return null;
    var q=s.match(/^['\"](.*)['\"]$/); if(q) return q[1];
    return vars()[s.replace(/^\$/,'')];
  }
  function applyOne(s){
    s=(s||'').trim(); if(!s) return;
    var m=s.match(/^\$?([A-Za-z_]\w*)\s*(\+=|-=|=)\s*(.*?)\s*$/); if(!m) return;
    var k=m[1], op=m[2], v=value(m[3]);
    if(op==='=') vars()[k]=v; else if(op==='+=') vars()[k]=(Number(vars()[k])||0)+Number(v||0); else vars()[k]=(Number(vars()[k])||0)-Number(v||0);
  }
  function apply(expr){ (expr||'').split(';').forEach(applyOne); }
  function truthy(expr){
    if(!expr) return true;
    var parts=expr.split(/\s*&&\s*/);
    for(var i=0;i<parts.length;i++){
      var p=parts[i].trim().replace(/^\$/,'');
      if(!vars()[p]) return false;
    }
    return true;
  }
  function feedback(type,text){
    if(!vars().scenarioFeedback) vars().scenarioFeedback=[];
    vars().scenarioFeedback.push({type:type,text:text});
  }
  function submitMulti(btn,target,spec,costRule,gate,gateFail){
    var passage=jQuery(btn).closest('.passage');
    var selected=[];
    passage.find('input.scenario-multi:checked').each(function(){
      var id=this.getAttribute('data-choice');
      for(var i=0;i<spec.length;i++) if(spec[i].id===id) selected.push(spec[i]);
    });
    if(!selected.length){ alert('Select at least one option before continuing.'); return false; }
    var costs=[];
    selected.forEach(function(c){ apply(c.onchoose); costs.push(Number(c.cost)||0); });
    var cost=0;
    if(costRule==='parallel-max-selected') cost=Math.max.apply(Math,costs.concat([0]));
    else { cost=costs.reduce(function(a,b){return a+b;},0); if(/\*\s*0\.75/.test(costRule)) cost=Math.ceil(cost*0.75); }
    vars().time=(Number(vars().time)||0)+cost;
    if(gate && !truthy(gate)){ state.display(gateFail || target, btn); return false; }
    state.display(target, btn); return false;
  }
  return {apply:apply,feedback:feedback,submitMulti:submitMulti};
})();
'''

EMBEDDED_SUGARCUBE_SHELL_ZLIB_B64 = """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"""

def embedded_template_html() -> str:
    import base64
    import zlib
    return zlib.decompress(base64.b64decode(EMBEDDED_SUGARCUBE_SHELL_ZLIB_B64)).decode("utf-8", errors="ignore")

EXTRA_CSS = r'''
/* ===== Scenario typography and spacing ===== */
#passages .passage{
    max-width:900px;
    margin:0 auto;
    padding:1.15rem 1.75rem;
    font-size:20px;
    line-height:1.55;
}
#passages .passage p{
    margin:0 0 .45rem 0;
}
#passages .passage h2{
    margin:0 0 .75rem 0;
    line-height:1.25;
}
#passages .passage h3{
    margin:.7rem 0 .25rem 0;
    line-height:1.3;
}
#passages .scenario-list{
    margin:.2rem 0 .55rem 1.2rem;
    padding-left:1rem;
}
#passages .scenario-list li{
    margin:0 0 .12rem 0;
    padding-left:.1rem;
    line-height:1.45;
}
.author-note{display:none}
.scenario-option{display:block;margin:.06rem 0;padding:.28rem .5rem;border:1px solid #d9d9d9;border-radius:5px;background:#fafafa;line-height:1.3}
.scenario-option input{margin:0 .4rem 0 0;transform:scale(1.02);vertical-align:middle}
.scenario-continue{margin:.45rem 0 0 0;font-size:1rem;padding:.45rem .8rem}
.scenario-action{margin:.18rem 0 0 0;padding:0;line-height:1.25}
.scenario-action .link-internal{margin:0;padding:.38rem .65rem}
.scenario-table{width:100%;border-collapse:collapse;margin:1rem 0}.scenario-table th,.scenario-table td{border:1px solid #cfcfcf;padding:.55rem;text-align:left}.scenario-table th{background:#f5f5f5}
.scenario-chart{height:170px;display:flex;align-items:flex-end;gap:5px;padding:10px;border-left:2px solid #555;border-bottom:2px solid #555;margin:1rem 0}.scenario-bar{flex:1;min-width:8px;background:#777}
.score-row{padding:.35rem 0;border-bottom:1px solid #ddd}
.link-internal{display:inline-block;margin:0 .3rem 0 0;padding:.38rem .65rem;border:1px solid #b7c9d6;border-radius:5px;background:#f7fbfe}
@media(max-width:800px){.passage{padding:1.25rem}.scenario-table{font-size:.9rem;display:block;overflow-x:auto}}
'''


def compile_html(story: Story) -> str:
    template = embedded_template_html()
    css = extract_custom_css(template) + "\n" + EXTRA_CSS
    passages_xml = []
    for idx,p in enumerate(story.passages, start=1):
        content = render_passage(p)
        passages_xml.append(f'<tw-passagedata pid="{idx}" name="{html.escape(p.title, quote=True)}" tags="" position="25,25" size="100,100">{html.escape(content)}</tw-passagedata>')
    start_pid = 1
    for idx,p in enumerate(story.passages, start=1):
        if p.title == story.start: start_pid=idx; break
    ifid = str(uuid.uuid4()).upper()
    storydata = (
        f'<tw-storydata name="{html.escape(story.title, quote=True)}" startnode="{start_pid}" creator="Word Scenario Converter" creator-version="0.1" '
        f'format="SugarCube" format-version="1.0.35" ifid="{ifid}" options="" tags="" zoom="1" hidden>'
        f'<style role="stylesheet" id="twine-user-stylesheet" type="text/twine-css">{css}</style>'
        f'<script role="script" id="twine-user-script" type="text/twine-javascript">{SCENARIO_JS}</script>'
        + ''.join(passages_xml) + '</tw-storydata>'
    )
    new = re.sub(r'<tw-storydata\b.*?</tw-storydata>', lambda _m: storydata, template, count=1, flags=re.S)
    new = re.sub(r'<title>.*?</title>', f'<title>{html.escape(story.title)}</title>', new, count=1, flags=re.S)
    legacy_title = "Literature Review: From Question to Synthesis"
    new = new.replace(legacy_title, story.title)
    if new == template:
        raise ValueError("Could not locate <tw-storydata> in SugarCube template HTML.")
    return new


def validate(story: Story) -> List[str]:
    titles = [p.title for p in story.passages]
    errors=[]
    if len(titles)!=len(set(titles)):
        errors.append("Duplicate Heading 2 passage names found.")
    if story.start not in titles:
        errors.append(f"StartPassage '{story.start}' is not a Heading 2 passage.")
    title_set=set(titles)
    for p in story.passages:
        body='\n'.join(x.text for x in p.items)
        for label,target in LINK_RE.findall(body):
            if target.strip() not in title_set:
                errors.append(f"Broken link in '{p.title}': target '{target.strip()}' does not exist.")
        for x in p.items:
            m=DIRECTIVE_RE.match(x.text)
            if m and m.group(1).lower() in {"choice", "multichoice"}:
                head,kv=parse_kv(m.group(2))
                target=kv.get('target')
                if not target and head and "->" in head:
                    target=head.split("->",1)[1].strip()
                if target and target not in title_set:
                    errors.append(f"Broken Choice target in '{p.title}': '{target}' does not exist.")
    return errors


def main():
    ap=argparse.ArgumentParser(description="Convert styled or directive-based Word scenarios to Twee and standalone SugarCube HTML.")
    ap.add_argument("input", type=Path)
    ap.add_argument("--output-dir", type=Path, default=Path("output"))
    ap.add_argument("--twee-only", action="store_true")
    args=ap.parse_args()
    story=load_docx(args.input)
    errors=validate(story)
    if errors:
        print("Validation failed:", file=sys.stderr)
        for e in errors: print(" -",e,file=sys.stderr)
        return 2
    args.output_dir.mkdir(parents=True, exist_ok=True)
    stem=args.input.stem
    twee_path=args.output_dir/(stem+".twee")
    twee_path.write_text(story_to_twee(story), encoding="utf-8")
    print(f"Wrote {twee_path}")
    if not args.twee_only:
        html_path=args.output_dir/(stem+".html")
        html_path.write_text(compile_html(story),encoding="utf-8")
        print(f"Wrote {html_path}")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
