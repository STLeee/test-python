import re

options = ["apple", "B. banana", "C. cat", "D. dog", "(a) ant", "1. one", "1.2"]

# Remove leading numbers, letters, and parentheses
pattern = r"^[\(a-dA-D0-4]+[\.\)]\s+"
options = [re.sub(pattern, "", option) for option in options]

print(options)
