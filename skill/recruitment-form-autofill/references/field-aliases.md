# Field aliases and semantic decisions

Use the label, surrounding section, units, help text, and available options together. Label similarity alone is insufficient. Aliases identify candidates, not interchangeable facts. Read record status, hints and notes; ask about contradictions or unclear applicability.

| Excel meaning | Common website labels/options | Safe action |
|---|---|---|
| 健康 | 健康状况=良好 | Fill and report normalization |
| 未参加 CPA | CPA通过情况=未考试 | Fill and report normalization |
| Exact monthly amount | Salary interval containing amount | Select containing interval and report |
| City name | Province/city hierarchy | Select narrowest available matching location |
| Numeric rank without cohort size | Percentage bands | Ask; do not infer percentage |

All actions above remain subject to SKILL.md authorization gates and profile-use, including confirmation for medical fields. These are generic meanings, not candidate facts.

## Repeatable sections / record splitting

When a form exposes an “添加/增加” control, treat the section as a list. Count
only clearly separate records in the Excel database, reuse the first visible
block, and add exactly `N-1` blocks for `N` records. Re-read the page after each
addition and verify one-to-one record mapping; do not duplicate blocks already
representing a record.

| Excel records | Website section | Mapping rule |
|---|---|---|
| One row per certificate or exam | 添加语言能力 | One credential per block; keep name and score paired |
| One row per school/degree | 添加教育经历 | Keep school, degree and dates from the same row |
| One row per employer/placement | 添加实习经历/工作经历 | Keep organization, role and dates together |
| One row per project/skill/honor | 添加项目/技能/荣誉 | One row becomes one block when the site supports it |

Do not split a free-text value merely because it contains commas, 顿号, or
semicolons. Ask when the database does not identify separate records or when a
split would alter facts, scores, dates, qualifications, or legal meaning.

## Decision criteria

- Exact values: preserve the database meaning and the requested record's scope.
- Equivalent option: select only with high confidence that the website's definition preserves the meaning. “未参加” may map to “未考试”, but never to “未通过” or “已报名” without evidence.
- Containing option: verify currency, monthly/annual period, tax basis and inclusive/exclusive interval endpoints. Select the unique interval containing the amount; do not select a merely nearby interval or invent a pay conversion. Overlap or missing units requires clarification.
- Hierarchical location: select the narrowest offered location that contains the known place. Do not infer district, street, relocation preference, household registration or birthplace from a city in another field. Ask if the city name or hierarchy is ambiguous.
- Imperfect but clearly best: proceed only if the option remains fact-preserving; record `Excel原值 → 网页选项 → 判断依据` with private values redacted in reports. Similar plausibility or changed qualifications, dates, scores, money, intent or legal meaning requires asking.
- Rank: a numeric position does not determine a percentage without cohort size and the website's band definition. Do not supply missing denominators.

## Common alias candidates

| Area | Candidate labels | Distinctions to preserve |
|---|---|---|
| Personal information | 姓名/中文姓名/法定姓名; 手机号码/联系电话; 邮箱/电子邮件 | Legal/display names and contact types may differ; 身份证号/证件号码 requires document type and confirmation. |
| Education | 毕业院校/学校; 专业/所学专业; 学历/最高学历; 学位; 入学时间/毕业时间 | Degree versus education level, current versus completed study, expected versus actual graduation. |
| Internship/work | 实习经历/工作经历; 单位/公司; 职位/岗位; 起止时间; 工作内容/职责 | Keep internship and employment categories separate; preserve organization, role and dates from the same record. |
| Language/certificates | 外语/语种; 语言水平/等级; 证书/资格; 考试成绩/通过情况 | Score versus level, registered versus attempted versus passed; never infer certification from a course. |
| Salary/location | 期望薪资/薪酬要求; 月薪/年薪; 意向城市/工作地点; 现居地/户籍地 | Check amount basis and interval boundaries; preference, current residence and household registration are separate facts. |
| Application intent | 意向岗位/申请职位; 到岗时间/可入职日期; 是否接受调剂/异地工作 | The target opening does not establish user preferences, availability or consent; ask if the database lacks them. |
