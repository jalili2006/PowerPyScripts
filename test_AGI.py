import matplotlib.pyplot as plt

# Data for the chart
labels = [
    "Global AI Market 2025 ($B)",
    "% Companies Using AI",
    "% Using Gen-AI",
    "% Scaled AI Agents",
    "% Productivity Gain Users",
    "AI Economic Impact 2030 ($B)"
]
values = [391, 78, 71, 2, 64, 15700]

# Create bar chart
plt.figure(figsize=(12, 6))
bars = plt.bar(labels, values, color=['#4c72b0', '#55a868', '#c44e52', '#8172b2', '#ccb974', '#64b5cd'])
plt.xticks(rotation=45, ha='right')
plt.title("Key Artificial Intelligence Statistics (2025 & Beyond)")
plt.ylabel("Values")
plt.grid(axis='y', linestyle='--', alpha=0.7)

# Annotate bars with values
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width() / 2, yval + max(values)*0.01, f'{yval}', ha='center', va='bottom')

plt.tight_layout()
plt.show()
