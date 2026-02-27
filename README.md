# kb.xlsx -> ICS

使用 `Python` 将 `kb.xlsx` 中的课程表转换为 `kb.ics`。

## 环境准备

```bash
uv sync
```

## 生成日历文件
![alt text](images/f7bb5306c262919e41731c77ba0b8161.jpg)
将新版mis系统首页的课程表复制粘贴到xlsx文件中即可食用。

```bash
uv run python src/kb_to_ics.py \
  --input kb.xlsx \
  --output kb.ics \
  --semester-monday 2026-03-02 \
  --timezone Asia/Shanghai
```
semester-monday后面填写学期的第一个周一的日期。

## 输出规则

- `SUMMARY`：`课程名(教室)`（如果教室为空，则使用 `未安排教室`）
- `DESCRIPTION` 第一行：`课程名 | 教室 | 教师`
- `DESCRIPTION` 还会保留班号、课程代码、周次范围、节次文本、星期和日期
- `LOCATION`：教室（如果 Excel 中有提供）
- 同一天同一课程会自动合并节次：`1,2 + 3,4 -> 1-4`、`6,7 + 8,9 -> 6-9`

## 当前节次时间映射

- `第1,2节` -> `08:30-10:10`
- `第3,4节` -> `10:30-12:10`
- `第6,7节` -> `14:25-16:05`
- `第8,9节` -> `16:25-18:05`

## 导入到日历

将 `kb.ics` 导入你的日历应用：

- Apple Calendar：`File -> Import`
- Google Calendar（网页）：`Settings -> Import & export -> Import`
- Outlook：`Add calendar -> Upload from file`

100% vibe coding. 仅供参考。
