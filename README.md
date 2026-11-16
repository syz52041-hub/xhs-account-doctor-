# 📊 xhs-account-doctor · 小红书账号诊断器

把小红书创作服务平台导出的「笔记列表明细表.xlsx」变成一份**能直接照着做的诊断看板**——不是摆数字，是找病根。

这是一个 [Agent Skills](https://agentskills.io) 开放标准格式的技能包：一个文件夹 + 一份 `SKILL.md`，装进支持该标准的 AI Agent（Claude Code、WorkBuddy 等）后，Agent 会在你提到"账号诊断""跑洞察"时自动按这套流程干活。

## 它解决什么问题

小博主最常见的困境：后台一堆数据，看不出问题在哪。这个工具自动做五件事：

1. **四段漏斗**（曝光→观看→互动→涨粉），一句话告诉你最大漏点在哪
2. **单篇红黑榜** + 黑榜三桶归因（没流量 / 有流量没接住 / 有效涨粉）
3. **爆款规律**：哪条内容线在涨粉效率上领跑、标题长度影响几何
4. **埋没检测**：互动率高但曝光 <100 的高质量笔记
5. **P0/P1/P2 优化清单**：每条都带数据依据 + 具体改法 + 保守预期

输出是一份**手机友好、零外部依赖的 HTML**，离线可看，数据不上传任何第三方。

## 特点

- **零爬取**：只用你自己后台导出的官方 xlsx，比第三方数据服务还准
- **零风险**：手动导入，不碰登录态、不踩封号灰产
- **不编造**：所有结论可追溯到真实数字，预期只做保守估算并标注
- **多导出自动合并去重**：计数取 max、点击率取最新快照、派生指标重算（踩坑修正过的口径）
- **自己 vs 对标**：`--compare` 一键并排对比，差距自动诊断

## 安装

### 方式一：npx 一键安装（推荐）

```bash
npx skills add <你的GitHub用户名>/xhs-account-doctor@xhs-account-doctor
```

### 方式二：手动复制

把 `xhs-account-doctor/` 文件夹整个复制到对应 Agent 的技能目录：

- Claude Code：`~/.claude/skills/xhs-account-doctor/`（全局）或项目的 `.claude/skills/`（单项目）
- WorkBuddy：`~/.workbuddy/skills/xhs-account-doctor/`
- GitHub Copilot：`~/.copilot/skills/xhs-account-doctor/`

## 使用

### 准备数据

小红书 App → 创作服务平台 → 数据中心 → 笔记数据 → 导出「笔记列表明细表.xlsx」。**多导出几份不同日期的**（每份只是切片，合并才全量）。

### 三种跑法

```bash
pip install pandas openpyxl        # 一次性依赖

# ① 自动扫描 ~/Downloads 里所有笔记导出表，合并诊断
python xhs-account-doctor/scripts/account_doctor.py

# ② 自己 vs 对标 并排对比
python xhs-account-doctor/scripts/account_doctor.py --compare 我的数据.xlsx 对标数据.xlsx

# ③ 自定义内容线打标（强烈建议：改成你自己账号的内容线）
cp xhs-account-doctor/scripts/config.example.json config.json
# 编辑 config.json 里的 tag_rules ...
python xhs-account-doctor/scripts/account_doctor.py --config config.json
```

### 让 AI 帮你跑（装了技能包后）

直接对 Agent 说：

> 帮我诊断一下小红书账号，导出表在 Downloads 里

Agent 会自动执行脚本、打开看板、并用一句话告诉你最大漏点在哪、P0 该干嘛。

## 内容线打标（重要）

每条笔记会按标题关键词自动归入你的"内容线"。默认规则（干货/热点）只是示例，**用之前先改成自己的**——比如美食号可以是「菜谱 / 探店 / 厨房好物」，母婴号可以是「辅食 / 遛娃 / 好物」。

## 样例看板

`docs/sample_dashboard.html` 是用合成数据生成的完整看板样例（数据为虚构演示，非真实账号）。

想先在真实安装前体验完整流程，可以生成演示数据自己跑一遍：

```bash
python xhs-account-doctor/scripts/make_demo_data.py    # 生成两份虚构的演示 xlsx
python xhs-account-doctor/scripts/account_doctor.py \
    xhs-account-doctor/scripts/demo_data/笔记列表明细表_0930.xlsx \
    xhs-account-doctor/scripts/demo_data/笔记列表明细表_1015.xlsx
```

## 方法学

数据口径与判定规则全部来自实战踩坑，写在 [`xhs-account-doctor/references/methodology.md`](xhs-account-doctor/references/methodology.md)：

- 为什么必须合并全部导出表（漏一份，结论倍数就错）
- 去重三类字段三种取法（点击率取 max 是错的，会留旧快照虚高值）
- 封面点击率是小数不是百分数
- 四段漏斗的断裂判定线（末端 <5%）

## 免责与边界

- 本工具只做**诊断分析**，不涉及爬虫、买卖数据、刷量等任何灰产
- 优化清单里的预期数字是保守估算，**不构成任何涨粉/效果承诺**
- 你的数据全程本地处理，产物 HTML 也在本地，不外传

## License

[MIT](LICENSE)
