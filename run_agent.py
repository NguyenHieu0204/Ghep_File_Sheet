import subprocess

def load_rules():
    with open("karpathy_rules.txt", "r", encoding="utf-8") as f:
        return f.read()

SYSTEM_RULES = load_rules()

def run_antigravity(user_input):
    full_prompt = f"""
[SYSTEM]
{SYSTEM_RULES}
[/SYSTEM]

[USER]
{user_input}
[/USER]
"""

    result = subprocess.run(
        ["antigravity", "run", full_prompt],
        capture_output=True,
        text=True
    )

    return result.stdout


if __name__ == "__main__":
    while True:
        user_input = input(">>> ")
        print(run_antigravity(user_input))