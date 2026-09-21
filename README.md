# agent-skills

个人维护的 AI Agent Skill 集合，从真实项目实践中蒸馏，即装即用。

## 安装

把整个仓库 clone 到 Agent 的 skill 目录（如 `~/.agents/skills/`）：

```bash
git clone git@github.com:Raylene-Y/agent-skills.git
cp -r agent-skills/<skill-name> ~/.agents/skills/
```

每个 Skill 为独立目录，含 `SKILL.md`（入口）、规则/模板/脚本/样例分层组织，详见各目录内说明。

## Skill 列表

| Skill | 说明 | 依赖 |
| --- | --- | --- |
| [pi-bid](pi-bid/) | 标书（投标技术方案）写作工作流：口径管理、评审规则清单、md→docx 导出管线、配图管线 | pandoc、.NET 9+ SDK（导出阶段）、soffice（配图阶段） |

## License

本仓库代码以 [MIT](LICENSE) 发布。`pi-bid/vendor/minimax-docx/` 为 MiniMaxAI 的第三方工具（MIT），其 LICENSE 随源码保留。
