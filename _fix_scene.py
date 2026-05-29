import re

with open("ai_test_script.py", "r", encoding="utf-8") as f:
    content = f.read()

# Locate old section
marker = "# ==================== C"
old_start = content.find(marker)
old_end = content.find("\ndef is_conversation_ended")
old_block = content[old_start:old_end]

with open("_scene_behaviors.py", "r", encoding="utf-8") as f:
    new_block = f.read()

content = content.replace(old_block, new_block)
with open("ai_test_script.py", "w", encoding="utf-8") as f:
    f.write(content)
print("Replaced OK")
