# -*- coding: utf-8 -*-
"""Парсер площадки НТИ (experts.nti.work, improject-47, TechNet): ищет свободные проекты под экспертизу.

Что делает:
  1. Собирает ленту TechNet со статусом «Не взято в экспертизу» (published).
  2. Собирает ленту «На экспертизе» (in_progress): эти проекты уже кто-то оценивает на площадке.
  3. Убирает проекты из таблицы распределения команды (любой статус) и из ленты «На экспертизе».
  4. Оценивает профиль TechNet по словам, открывает страницу проекта на pt.2035:
     ищет презентацию и стадию.
  Готовые экспертизы: на карточке есть счётчик «Экспертные заключения: N», при N > 0 проект отсеивается.
  5. Пишет candidates.csv (лучшие сверху) и all_projects.json (вся лента).

Запуск:
  pip install requests
  python nti_parser.py --plan plan.csv --pages 40 --top 60
plan.csv — таблица распределения, выгруженная в CSV (колонка с ссылками ideas/NNNNN). Можно .xlsx (нужен openpyxl).
"""
import argparse, csv, html, json, re, sys, time
import requests

BASE = "https://experts.nti.work/improject-47"
FEED = BASE + "?expertise_status%5B%5D={st}&field1624555270698=126&kept_view_type=ideas&order=new&page={p}"
UA = {"User-Agent": "Mozilla/5.0"}

POS = (r"станк|производств|металл|детал|оборудован|сварк|лить|литей|аддитив|3d|композит|датчик|контрол|чпу|"
       r"двигател|робот|промышлен|автоматиз|материал|покрыт|подшипник|турбин|инструмент|машиностро|прокат|"
       r"нефтепер|завод|станоч|фрез|резц|сплав|порош|изготовл|диагност|износ|вибрац")
NEG = (r"школ|студент|образоват|ферм|агро|урожа|растен|животн|медиц|пациент|здоров|туризм|кафе|ресторан|"
       r"игр[аы]|дрон|бпла|беспилот|финанс|магазин|маркетплейс|мероприят|чат-бот|телеграм")


def get(url):
    for _ in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=30)
            if r.ok:
                return r.text
        except requests.RequestException:
            pass
        time.sleep(2)
    return ""


def clean(s):
    return html.unescape(re.sub(r"<[^>]+>", " ", s or "")).strip()


def parse_feed(h):
    out = []
    for m in re.finditer(r'data-id="(\d+)" data-show-in="feed" id="article-\d+">(.*?)article__footer', h, re.S):
        i, b = m.group(1), m.group(2)
        t = re.search(r"article__title[^>]*><span>(.*?)</span>", b, re.S)
        d = re.search(r'</h2><div class="article__text">(.*?)</div>', b, re.S)
        f = {clean(k): clean(v) for k, v in re.findall(
            r'<h3 class="article__subtitle[^>]*>(.*?)</h3><div class="article__text">(.*?)</div>', b, re.S)}
        link = re.search(r'href="(//pt\.2035[^"]+|https?://[^"]+)" *>На страницу', b)
        link = link.group(1) if link else ""
        if link.startswith("//"):
            link = "https:" + link
        ex = re.search(r"Экспертные заключения:(?:&nbsp;|\s)*(\d+)", b)
        out.append(dict(id=i, title=clean(t.group(1) if t else ""), desc=clean(d.group(1) if d else ""),
                        expertises=int(ex.group(1)) if ex else -1,
                        dir=f.get("Направление", ""), tech=f.get("Сквозные технологии", ""),
                        markets=f.get("Рынки", ""), pt=link, url=f"{BASE}/ideas/{i}"))
    return out


def crawl(status, pages):
    res = []
    for p in range(1, pages + 1):
        items = parse_feed(get(FEED.format(st=status, p=p)))
        if not items:
            break
        res += items
        print(f"  {status} стр. {p}: {len(items)}", file=sys.stderr)
        time.sleep(0.4)
    return res


def plan_ids(path):
    if not path:
        return set(), {}
    rows = []
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        for ws in openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets:
            rows += [[str(c) if c is not None else "" for c in r] for r in ws.iter_rows(values_only=True)]
    else:
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
    ids, info = set(), {}
    for r in rows:
        line = " | ".join(r)
        for m in re.finditer(r"ideas/(\d+)", line):
            ids.add(m.group(1)); info[m.group(1)] = line[:200]
    return ids, info


def project_page(url):
    """Презентация и стадия со страницы pt.2035."""
    h = get(url) if url else ""
    pres = re.findall(r'https://s3[^"\']+/presentation/[^"\']+\.(?:pptx?|pdf)', h)
    stage = re.search(r"(Идея или концепция|Прототип[^<\"]{0,30}|MVP[^<\"]{0,20}|Опытный образец[^<\"]{0,20}|"
                      r"Продажи[^<\"]{0,20}|Масштабирование[^<\"]{0,20}|Готовый продукт[^<\"]{0,20})", h)
    return dict(pres=pres[0] if pres else "", stage=stage.group(1) if stage else "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", help="таблица распределения команды (csv/xlsx)")
    ap.add_argument("--pages", type=int, default=40, help="сколько страниц ленты смотреть")
    ap.add_argument("--top", type=int, default=60, help="сколько лучших открыть на pt.2035")
    ap.add_argument("--min-score", type=int, default=4)
    a = ap.parse_args()

    taken, info = plan_ids(a.plan)
    print(f"В таблице команды: {len(taken)} проектов", file=sys.stderr)
    free = crawl("published", a.pages)
    busy_items = crawl("in_progress", a.pages)
    busy = {x["id"] for x in busy_items}
    print(f"Лента «Не взято»: {len(free)}, «На экспертизе»: {len(busy)}", file=sys.stderr)
    json.dump(dict(free=free, in_progress=sorted(busy)), open("all_projects.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # если одна из карточек-дублей занята, заняты все
    seen = {x["pt"] or x["title"].lower() for x in free
            if x["id"] in taken or x["id"] in busy or x["expertises"] > 0}
    seen |= {x["pt"] or x["title"].lower() for x in busy_items}
    cand = []
    for x in free:
        key = x["pt"] or x["title"].lower()
        if key in seen:
            continue  # один проект бывает заведён несколькими карточками
        seen.add(key)
        if x["id"] in taken or x["id"] in busy or x["expertises"] > 0:
            continue  # взят командой, на экспертизе на площадке или уже есть экспертные заключения
        t = " ".join([x["title"], x["desc"], x["tech"], x["dir"]]).lower()
        x["score"] = len(re.findall(POS, t)) - 3 * len(re.findall(NEG, t)) + min(len(x["desc"]) // 150, 4)
        if x["score"] >= a.min_score:
            cand.append(x)
    cand.sort(key=lambda x: -x["score"])

    rows = []
    for x in cand[:a.top]:
        x.update(project_page(x["pt"]))
        rows.append(x)
        time.sleep(0.3)
    rows.sort(key=lambda x: (not x["pres"], -x["score"]))

    cols = ["id", "score", "stage", "pres", "expertises", "title", "markets", "dir", "url", "pt", "desc"]
    with open("candidates.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore"); w.writeheader(); w.writerows(rows)
    ok = [x for x in rows if x["pres"]]
    print(f"Кандидатов: {len(rows)}, с презентацией: {len(ok)} -> candidates.csv")
    for x in ok[:10]:
        print(x["id"], x["score"], x["stage"] or "-", "|", x["title"][:90])


if __name__ == "__main__":
    main()
