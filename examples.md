# 示例

示例一是用户确认过、可直接交给执行 Agent 的完整输出，保持原样。示例二只示范信息不足时怎么问。

## 示例一：完整指令

**不要解析校招页 HTML。** 京东 2027 校招在 `campus.jd.com`，公开接口就能按项目、方向、城市筛。

入口：`https://campus.jd.com/`（2027 校招进行中）  
职位大厅：`https://campus.jd.com/#/jobs`  
应届列表：`POST https://campus.jd.com/api/wx/position/page?type=present`  
`pageIndex` 从 **0** 开始。项目/方向先调 `GET .../api/wx/position/getProjectList`，对上 **JDS / 新星计划**。详情：`GET .../api/wx/position/detail/{publishId}`。

字段大致对应：`publishId` → 职位ID，项目名（新星计划等）→ 性质，`jobDirection` → 类型，`positionName` → 职位名称，`workContent` → 职位描述，`qualification` → 职位要求。城市在 `requirementVoList[].workCity` / `workCityCode`。

方向码以 `getProjectList` 实时为准，历史上常见：`03` 产品、`04` 运营、`17` 市场（投放多半在这）。**排除**技术 `02`、TGT `14`、职能 `34` 以及人力/财务/法务。城市码不要写死，从接口里对上上海、深圳、杭州、北京。

---

复制下面整段，在 **Agent 模式**、项目文件夹里发出去：

```text
帮我做一个京东 2027 校招信息采集脚本，输出 Excel。

目标站点（只要校招 campus.jd.com，不要社招 zhaopin.jd.com）：
- https://campus.jd.com/
- 职位大厅：https://campus.jd.com/#/jobs
- 2027 应届主投：招聘职位里的 JDS / 新星计划（type=present）
不要爬 HTML。先试请求确认字段和筛选项，再写脚本。

公开接口：
1. GET https://campus.jd.com/api/wx/position/getProjectList
   确认 2027 应届项目（JDS/新星计划）的 planId、职位方向码、城市码。
2. POST https://campus.jd.com/api/wx/position/page?type=present
   body 形态（以实际抓包为准，pageIndex 是 0 起）：
   {
     "pageSize": 20,
     "pageIndex": 0,
     "parameter": {
       "positionName": "",
       "planIdList": [<新星计划 planId>],
       "positionDeptList": [],
       "jobDirectionCodeList": [<产品/运营/市场等方向码>],
       "workCityCodeList": [<上海/深圳/杭州/北京的城市码>]
     }
   }
3. 列表里 workContent / qualification 不完整时，再调
   GET https://campus.jd.com/api/wx/position/detail/{publishId}

请求头带：
- Origin: https://campus.jd.com
- Referer: https://campus.jd.com/
- Content-Type: application/json
- 普通浏览器 User-Agent

范围（必须同时满足）：
- 批次：2027 校招应届，优先 JDS/新星计划。不要实习生计划、不要 TGT 顶尖技术、不要社招。
- 城市（优先顺序，只保留办公地包含这些城市的岗位；多城市岗位只要命中任一城市就保留）：
  1. 上海  2. 深圳  3. 杭州  4. 北京
  城市码从 getProjectList 或列表 requirementVoList 反查，不要写死猜测。
- 职位方向：产品、运营、投放/市场/商务，以及其他非技术、非职能的业务岗。
  建议纳入：产品、运营、市场、投放、采销、销售、商务、设计、供应链/物流业务类、健康等业务方向。
  必须排除：技术/研发/算法/客户端/前端/后端/测开/数据研发、TGT、工程开发、职能（人力/财务/法务/行政等）、管培生（除非名称明显是产品运营类）。
  方向码以 getProjectList 为准；接口筛不干净时，再用职位名称+方向中文做二次过滤。
- 「投放」可能挂在市场/运营/广告下，用关键词（投放、广告、效果、信息流、媒介）补捞，但仍排除技术岗。

输出文件：jd_2027_campus.xlsx
Sheet 列（表头必须完全一致，这 6 列不要改名）：
职位ID	性质	类型	职位名称	职位描述	职位要求

字段映射：
- 职位ID：publishId
- 性质：应届/新星计划等招聘项目名（present → 校招应届）
- 类型：jobDirection 中文（如 产品方向、运营方向）；能区分投放/市场就写更细
- 职位名称：positionName
- 职位描述：workContent
- 职位要求：qualification
城市只做筛选，不要新增列。若同一 publishId 多城市，去重后保留一行。

要求：
1. Python + requests + openpyxl（或 pandas）
2. 分页拉取，pageIndex 从 0 递增，用 totalNumber 判断结束，按 publishId 去重
3. 请求间隔 1–2 秒；pageSize 用 20；失败重试
4. success 不为 true 时打印 errorMessage
5. 描述和要求保留换行，Excel 可读
6. 跑完打印：总条数、各城市命中数、各方向条数；xlsx 放到项目根目录
7. 先拉 1–2 页核对六个字段、城市和方向过滤是否正确，再全量
8. 不要登录京东账号，不要打需要 SSO/passport 的投递接口，不要绕过验证码/WAF

先试请求确认 2027 新星计划 planId、产品/运营/市场方向码、四城 workCityCode，再写脚本。
```

发之前先打开一个空项目目录（例如 `~/Projects/jd-2027-jobs`），再在那个工作区里下这条指令。

## 示例二：信息不足时先问

用户：「帮我爬校招 JD。」

```text
先确认这几项，我再写一份可交给任意 AI 执行的采集指令：

1. 哪家公司？只要校招，还是实习/社招也要？有官网链接更好。
2. 哪一届？只要正式应届，还是实习也收录？
3. 城市优先级？例如上海 > 深圳 > 杭州 > 北京。多城市岗位命中一个就保留吗？
4. 想要哪些岗位，明确不要哪些？例如要产品/运营/投放，不要技术和职能。
5. 这次只要指令，还是让我直接采集并导出 Excel？
```
