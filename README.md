# nti-expertizy – скил Claude Code для экспертиз TechNet (НТИ)

Находит свободные проекты TechNet на experts.nti.work (без готовых экспертиз и не занятые командой), помогает написать экспертизу по формату с критериями TRL/MRL/CRL/IRL, оформляет docx/html и проверяет.

## Установка

```bash
git clone https://github.com/<владелец>/nti-expertizy-skill.git
```

Скопировать папку `nti-expertizy` в `~/.claude/skills/` (Windows: `C:\Users\<имя>\.claude\skills\`), затем:

```bash
pip install requests python-docx openpyxl
```

## Использование

В Claude Code: «напиши 3 экспертизы НТИ, исполнитель <Имя>» или `/nti-expertizy`. Claude попросит таблицу распределения команды и дальше пойдёт по шагам из `SKILL.md`.
