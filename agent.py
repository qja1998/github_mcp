import asyncio
import shutil

import os
import dotenv
dotenv.load_dotenv(override=True)
ROOT = os.path.abspath(os.path.dirname(__file__))

from collections import deque

GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")

from agents import Agent, Runner, trace, WebSearchTool
from agents.mcp import MCPServer, MCPServerStdio


async def run(mcp_server: MCPServer, directory_path: str):
    agent = Agent(
        model=OPENAI_MODEL,
        name="Assistant",
        instructions=f"Answer questions about the git repositories, use that for repo_path",
        mcp_servers=[mcp_server],
        tools=[WebSearchTool()]
    )
    

    PROMPT = f"""
    주어진 회사에 대한 정보를 제공하는 web search assistant입니다.
    """

    chat_history = deque([], maxlen=5)

    while True:
        # Ask the user for the git command
        command = input("Please enter a git command (or 'exit' to quit): ")
        
        if command.lower() == "exit":
            break
        
        END_PROMPT = f"""
        history: {chat_history}
        user: {command}"""

        prompt = PROMPT + END_PROMPT
        print(prompt)

        # Run the command and print the result
        print("\n" + "-" * 40)
        print(f"Running: {command}")
        result = await Runner.run(starting_agent=agent, input=prompt)
        print(result.final_output)
        current_chat = [
            {"role": "user", "content": command},
            {"role": "assistant", "content": result.final_output},
        ]
        chat_history.append(current_chat)
        print("\n" + "-" * 40)


async def main():
    # Ask the user for the directory path
    directory_path = "https://github.com/qja1998"

    async with MCPServerStdio(
        cache_tools_list=True,  # Cache the tools list, for demonstration
        params={
            # "command": "docker",
            # "args": [
            #     "run",
            #     "-i",
            #     "--rm",
            #     "-e",
            #     "GITHUB_PERSONAL_ACCESS_TOKEN",
            #     "ghcr.io/github/github-mcp-server"
            # ],
            # "env": {
            #     "GITHUB_PERSONAL_ACCESS_TOKEN": GITHUB_PERSONAL_ACCESS_TOKEN,
            # }

            "command": "npx",
            "args": [
                "-y",
                "@modelcontextprotocol/server-sequential-thinking"
            ]

            
        },
    ) as server:
        with trace(workflow_name="MCP Git Example"):
            await run(server, directory_path)


if __name__ == "__main__":
    if not shutil.which("uvx"):
        raise RuntimeError("uvx is not installed. Please install it with `pip install uvx`.")

    asyncio.run(main())