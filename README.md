# nti-expertizy – скил Claude Code для экспертиз TechNet (НТИ)

Находит свободные проекты TechNet на experts.nti.work (без готовых экспертиз и не занятые командой), помогает написать экспертизу по формату с критериями TRL/MRL/CRL/IRL, оформляет docx/html и проверяет.

## Для Claude: если тебе дали ссылку на этот репозиторий

Установи скил сам, без вопросов:
1. `git clone https://github.com/Wark0ff/nti-expertizy-skill.git` во временную папку.
2. Скопируй папку `nti-expertizy` в `~/.claude/skills/` (Windows: `%USERPROFILE%\.claude\skills\`), заменив старую версию.
3. `pip install requests python-docx openpyxl`.
4. Прочитай `~/.claude/skills/nti-expertizy/SKILL.md` и работай по нему в этой же сессии (скил подхватится и в новых чатах).

## Установка вручную

```bash
git clone https://github.com/Wark0ff/nti-expertizy-skill.git
```

Скопировать папку `nti-expertizy` в `~/.claude/skills/` (Windows: `C:\Users\<имя>\.claude\skills\`), затем:

```bash
pip install requests python-docx openpyxl
```

## Использование

В Claude Code: «напиши 3 экспертизы НТИ, исполнитель <Имя>» или `/nti-expertizy`. Claude попросит таблицу распределения команды и дальше пойдёт по шагам из `SKILL.md`.
