# -*- coding: utf-8 -*-
"""Оформление экспертиз НТИ в .docx и .html + проверка критериев на дословность.

Вход: JSON-файл со списком экспертиз:
[
  {"short": "Короткое имя", "title": "Полное название проекта",
   "paras": ["Абзац 1", "Абзац 2", "Абзац 3"],
   "criteria": ["TRL-3. ...", "MRL-1. ...", "CRL-3. ...", "IRL-0. ...", "Новизна ...", "Влияет", "Другое"],
   "id": "76045"}
]
Запуск: python mk_exp.py expertizy.json --who Имя --out папка
"""
import argparse, datetime, json, os, re, sys
from docx import Document
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
C = ["Уровень готовности технологии (TRL)", "Уровень готовности к рынку (MRL)", "Уровень команды (CRL)",
     "Уровень инвестиционной привлекательности (IRL)", "Новизна идеи (актуальность)",
     "Влияние на достижение/закрытие технологического суверенитета",
     "Приоритетное направление НТИ, к которому относится проект"]
URL = "https://experts.nti.work/improject-47/ideas/"


def allowed():
    """Допустимые значения каждого критерия из references/criteria.md."""
    txt = open(os.path.join(HERE, "..", "references", "criteria.md"), encoding="utf-8").read()
    txt = txt.split("## Образец")[0]
    res, cur = {}, None
    for line in txt.splitlines():
        m = re.match(r"\*\*(.+)\*\*$", line.strip())
        if m and m.group(1) in C:
            cur = m.group(1); res[cur] = []
        elif cur and line.startswith("- "):
            res[cur].append(line[2:].strip())
    return res


def check(e, ok):
    err = []
    if len(e["paras"]) != 3: err.append("нужно 3 абзаца")
    if len(e["criteria"]) != 7: err.append("нужно 7 значений критериев")
    for c, v in zip(C, e["criteria"]):
        if v.strip() not in ok.get(c, []): err.append(f"«{c}»: значение не совпадает дословно: {v[:60]}")
    body = " ".join(e["paras"])
    if "—" in body: err.append("длинное тире «—» в тексте, нужно «–»")
    if re.search(r"[\w.+-]+@[\w-]+\.\w+|\+7[\d\s()-]{9,}|8\s?\(?\d{3}", body): err.append("похоже на контакты")
    if not re.fullmatch(r"\d+", str(e["id"])): err.append("ID проекта не число")
    return err


def fill(d, e):
    d.add_paragraph().add_run(e["title"]).bold = True
    for t in e["paras"]: d.add_paragraph(t)
    for c, v in zip(C, e["criteria"]):
        p = d.add_paragraph(); p.add_run(c).bold = True; p.add_run("\n" + v)
    d.add_paragraph("Проект: " + URL + str(e["id"]))


def new_doc():
    d = Document(); st = d.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(11)
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json"); ap.add_argument("--who", default="Эксперт"); ap.add_argument("--out", default=".")
    a = ap.parse_args()
    E = json.load(open(a.json, encoding="utf-8"))
    ok = allowed(); bad = False
    for e in E:
        for m in check(e, ok):
            print(f"[{e.get('short')}] {m}"); bad = True
    if bad:
        sys.exit("Исправь ошибки, файлы не созданы.")
    os.makedirs(a.out, exist_ok=True)
    allin = new_doc()
    for i, e in enumerate(E):
        d = new_doc(); fill(d, e); d.save(os.path.join(a.out, f"Экспертиза_{e['short']}.docx"))
        if i: allin.add_page_break()
        fill(allin, e)
        h = "<html><body style='font-family:Arial;font-size:11pt'><p><b>%s</b></p>%s%s<p>Проект: %s%s</p></body></html>" % (
            e["title"], "".join(f"<p>{t}</p>" for t in e["paras"]),
            "".join(f"<p><b>{c}</b><br>{v}</p>" for c, v in zip(C, e["criteria"])), URL, e["id"])
        open(os.path.join(a.out, f"Экспертиза_{e['short']}.html"), "w", encoding="utf-8").write(h)
    allin.save(os.path.join(a.out, f"{datetime.date.today():%Y_%m%d}_Экспертизы_{a.who}.docx"))
    print("готово:", len(E), "->", os.path.abspath(a.out))


if __name__ == "__main__":
    main()
