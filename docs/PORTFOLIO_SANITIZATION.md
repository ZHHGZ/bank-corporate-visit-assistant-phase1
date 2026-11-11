# 作品集副本整理记录

- 来源：现有 `github/master` 一期代码快照，HEAD 为 `e6ad322`（2026-07-29）。
- 历史边界：可核实的一期 Git 历史从 2026-07-22 开始；没有证据支持 4 月中至 5 月末的提交历史。本副本不伪造或回填提交时间，也不携带未经清理的完整历史。
- 脱敏：机构名称、ESB 调用标识和服务身份字段已匿名化；端点改为 `example.invalid`；示例环境变量清空或设为演示值；已排除本地工具配置、部署编排、原始 DOCX 报告及原始项目文档。
- 范围：保留一期 API、服务编排、ESB 适配边界、解析、聚合、报告和测试代码。二期 Agent 代码不在副本中。
- 测试：Python 3.10.11；`PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 py -3.10 -m pytest tests/ -q --ignore=tests/test_formal_report_parity.py`，结果为 74 passed、1 skipped、1 failed。唯一失败是 `test_report_field_mapping.py::test_validation_passes` 对 `tech_qualifications` 映射缺少 `source_arrays` 和 `notes` 的配置校验。该字段缺失在原始快照和脱敏副本中都存在；没有为公开副本修改业务逻辑或掩盖失败。
- 发布状态：这是待人工审阅的本地候选，不代表组织已批准公开源代码或业务配置。