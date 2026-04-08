# ClaudeCode

A Python project for experimenting with LLM integrations using [LangChain](https://python.langchain.com/) and the [Anthropic Claude](https://www.anthropic.com/) API.

## Overview

This project demonstrates:
- Making LLM calls via `langchain-anthropic` (Claude models)
- Working with pandas DataFrames
- Managing API keys with `python-dotenv`

## Project Structure

```
ClaudeCode/
├── 1_llm_call.ipynb      # Notebook: basic LLM call with Claude via LangChain
├── create_dataframe.py   # Script: create a 10×5 pandas DataFrame
├── main.py               # Entry point
├── pyproject.toml        # Project metadata and dependencies
├── .env                  # API keys (not committed)
└── .gitignore
```

## Requirements

- Python >= 3.11
- [uv](https://github.com/astral-sh/uv) (recommended package manager)

## Setup

1. **Clone the repo and install dependencies:**

   ```bash
   uv sync
   ```

2. **Set up your environment variables:**

   Create a `.env` file in the project root:

   ```env
   ANTHROPIC_API_KEY=your_api_key_here
   ```

## Usage

**Run the main entry point:**

```bash
uv run main.py
```

**Create a sample DataFrame:**

```bash
uv run create_dataframe.py
```

**Explore the LLM notebook:**

Open `1_llm_call.ipynb` in Jupyter or VS Code to see an example of invoking Claude via LangChain.

## Dependencies

| Package               | Purpose                          |
|-----------------------|----------------------------------|
| `langchain-anthropic` | LangChain integration for Claude |
| `dotenv`              | Load environment variables       |
| `pandas`              | DataFrame manipulation           |
| `numpy`               | Numerical operations             |
