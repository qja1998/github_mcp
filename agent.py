from openai import Agent
from openai.mcp import MCPServerStdio, MCPServerSse


# MCP 서버 설정 (로컬 및 원격)
# mcp_server_1 = MCPServerStdio(params={"command": "npx", "args": ["@modelcontextprotocol/server-filesystem"]})
mcp_server_2 = MCPServerSse(url="https://ghcr.io/github/github-mcp-server")

# OpenAI Agents SDK에서 MCP 서버 활용
agent = Agent(
    name="Assistant",
    instructions="Use the tools to achieve the task",
    mcp_servers=[mcp_server_2]
)

print("Agnet tool: ", agent.tools)
print("Agent resources: ", agent.resources)

# agent.run()