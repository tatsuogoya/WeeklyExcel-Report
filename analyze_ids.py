import re

ids = [
    "SCTASK0980440", "SCTASK0979667", "SCTASK0977438", "SCTASK0981107",
    "SCTASK0981116", "SCTASK0981127", "SCTASK0981135", "SCTASK0981142",
    "SCTASK0981152", "SCTASK0981641", "SCTASK0981695", "SCTASK0981779",
    "SCTASK0981765", "SCTASK0981787", "SCTASK0981825", "SCTASK0981856",
    "SCTASK0981855", "SCTASK0979040", "SCTASK0978291", "SCTASK0981964",
    "SCTASK0981949", "SCTASK0981696", "SCTASK0979042", "SCTASK0981826",
    "SCTASK0982551", "INC19774449",   "SCTASK0982554", "SCTASK0983054",
    "SCTASK0983055", "SCTASK0983098", "SCTASK0976928", "SCTASK0983236",
    "SCTASK0972368", "SCTASK0971136", "SCTASK0983141", "SCTASK0983243",
    "SCTASK0979041"
]

# Separate by prefix
sctasks = []
incs = []

for i in ids:
    if i.startswith("SCTASK"):
        sctasks.append(int(i.replace("SCTASK", "")))
    elif i.startswith("INC"):
        incs.append(int(i.replace("INC", "")))

sctasks.sort()
incs.sort()

print(f"Total SCTASKs: {len(sctasks)}")
print(f"Total INCs: {len(incs)}")

print("\n--- SCTASK Analysis ---")
print(f"Min: {min(sctasks)}")
print(f"Max: {max(sctasks)}")

print("\n--- Gaps in SCTASK sequence (checking small gaps within clusters) ---")
# Finding gaps in a large range might be too noisy if the IDs are sparse.
# Let's check for "expected" sequences.
# Usually, if we see 981107, 981116... 
# If we assume they are sequential, there are HUGE gaps.
# User question: "What data are missing from this?"
# Maybe I should just check if there are any obvious "1 missing in a sequence of 10".

# Let's print the sorted list to see patterns
for s in sctasks:
    print(s)

# Check for consecutive runs
print("\n--- Runs ---")
prev = sctasks[0]
run = [prev]
for x in sctasks[1:]:
    if x == prev + 1:
        run.append(x)
    else:
        if len(run) > 1:
            print(f"Run: {run[0]} - {run[-1]}")
        else:
            print(f"Single: {run[0]}")
        run = [x]
    prev = x
if len(run) > 1:
    print(f"Run: {run[0]} - {run[-1]}")
else:
    print(f"Single: {run[0]}")
