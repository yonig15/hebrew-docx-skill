<div dir="rtl">

# מסמכי Word בעברית, מימין לשמאל

סקיל ל-Claude שמייצר קבצי Word בעברית שנפתחים נכון ב-Word, ב-Pages וב-Google Docs: פסקאות וטבלאות מימין לשמאל, פיסוק בצד הנכון, וטווחי מספרים כמו "5%-15%" שלא מתהפכים. הוא יודע גם לתקן קובץ Word בעברית שיצא הפוך.

## התקנה ב-Claude (אתר או אפליקציית דסקטופ)

לפני הכול: **Settings ← Capabilities**, ולהדליק את **Code execution and file creation**.

### דרך 1: ישירות מהריפו (מנוי בתשלום)

1. **Customize ← Plugins**.
2. באזור **Personal plugins**: הפלוס ← **Add marketplace** ← **Add from a repository**.
3. להדביק את הכתובת: `https://github.com/yonig15/hebrew-docx-skill`
4. להתקין את הפלאגין **hebrew-docx**.

### דרך 2: קובץ ZIP (עובד גם בתוכנית החינמית)

1. להוריד את [hebrew-docx.zip](https://github.com/yonig15/hebrew-docx-skill/releases/latest/download/hebrew-docx.zip).
2. **Customize ← Skills** ← הפלוס ← **Create skill** ← **Upload a skill**, לבחור את הקובץ ולהדליק את המתג.

### Claude Code

```
/plugin marketplace add yonig15/hebrew-docx-skill
/plugin install hebrew-docx@hebrew-docx-skill
```

## שימוש

פשוט לבקש, למשל: "תכין לי מסמך Word עם האפיון". כדי לתקן קובץ קיים: להעלות אותו ולכתוב "תתקן את הכיוון".

## מה בפנים

- `skills/hebrew-docx/SKILL.md` - ההוראות ל-Claude
- `skills/hebrew-docx/scripts/hebrew_docx.py` - הממיר. הרצה עם `--self-test` בודקת שהוא עובד.

נבדק ב-Microsoft Word: מסמך חדש, ותיקון של מסמך עברי שנוצר בכלי AI אחר.

</div>
