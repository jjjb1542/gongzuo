# Skills

本仓库收录自用的 AI Agent Skill（每个 Skill 一个子目录，入口均为 `SKILL.md`）。

| Skill | 用途 | 入口 |
|---|---|---|
| `asset-inventory-checktable` | 资产清查检查程序表：按 9 项清查程序、5 张表结构生成带「符合／部分符合／不符合」结论的 Excel 底稿 | [SKILL.md](asset-inventory-checktable/SKILL.md) |
| `rent-audit-workingpaper` | 房产（土地）租金收缴专项核查底稿：按 9 条审计程序生成 5 张表（资料清单、程序表、欠租台账、对账差异、疑点清单） | [SKILL.md](rent-audit-workingpaper/SKILL.md) |

## 目录结构

```
skills/
├─ asset-inventory-checktable/
│  ├─ SKILL.md
│  ├─ assets/        # Excel 模板与示例
│  ├─ references/    # 口径、程序、表结构说明
│  └─ scripts/       # 生成与校验脚本
└─ rent-audit-workingpaper/
   ├─ SKILL.md
   ├─ assets/
   ├─ references/
   └─ scripts/
```

## 使用说明

- `SKILL.md` 的 frontmatter（`name` / `description`）为触发依据，正文为执行步骤与口径。
- 脚本依赖 Python 3 与 `openpyxl`；模板、示例文件已随仓库提供。
- 生成的底稿属于业务产出，建议输出到仓库外目录，避免把交付物提交进版本库。
