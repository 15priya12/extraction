from spire.doc import *
import os, shutil, re

input_file = "demo2.docx"
temp_file = "temp_copy.docx"
output_md = "output_table.md"

# Safe file copy
shutil.copy(input_file, temp_file)

doc = Document()
doc.LoadFromFile(temp_file, FileFormat.Auto)


# --- Helper functions ---
def get_bullet_level(bullet):
    if not bullet:
        return {"level": 0, "type": "none", "value": ""}

    clean = bullet.strip().rstrip(".")

    if re.match(r"^(i{1,3}|iv|v|vi{0,3}|ix|x)$", clean, re.I):
        return {"level": 3, "type": "roman", "value": clean}
    elif re.match(r"^[a-z]$", clean):
        return {"level": 2, "type": "letter", "value": clean}
    elif re.match(r"^[A-Z]$", clean):
        return {"level": 2, "type": "letter-upper", "value": clean}
    elif re.match(r"^\d+$", clean):
        return {"level": 1, "type": "number", "value": clean}
    elif re.match(r"^\d+\.\d+\.\d+$", clean):
        return {"level": 3, "type": "number-dot-dot", "value": clean}
    elif re.match(r"^\d+\.\d+$", clean):
        return {"level": 2, "type": "number-dot", "value": clean}

    return {"level": 0, "type": "unknown", "value": clean}


def build_hierarchical_number(bullet_info, hierarchical_levels):
    """Keeps exact bullet label hierarchy, properly resetting when needed."""
    if not bullet_info or bullet_info["level"] == 0:
        return ""

    lvl = bullet_info["level"]

    # Reset deeper levels
    hierarchical_levels[:] = hierarchical_levels[:lvl]

    # Insert or update current level
    if len(hierarchical_levels) < lvl:
        hierarchical_levels.append(bullet_info["value"])
    else:
        hierarchical_levels[lvl - 1] = bullet_info["value"]

    # Rebuild full path
    return ".".join(hierarchical_levels).lstrip(".")


# --- Main logic ---
rows_data = []
hierarchical_levels = []

for s_idx in range(doc.Sections.Count):
    section = doc.Sections.get_Item(s_idx)

    for p_idx in range(section.Paragraphs.Count):
        para = section.Paragraphs.get_Item(p_idx)
        text = para.Text.strip()
        if not text:
            continue

        bullet_text = ""
        if para.ListFormat.ListType != ListType.NoList:
            raw_bullet = para.ListText.strip()
            info = get_bullet_level(raw_bullet)
            bullet_text = build_hierarchical_number(info, hierarchical_levels)
        else:
            hierarchical_levels.clear()

        rows_data.append((text, bullet_text))

doc.Close()
os.remove(temp_file)

# --- Markdown export ---
md_lines = [
    "| Paragraph Text | Bullet/List Label |",
    "|----------------|------------------|"
]
for text, bullet in rows_data:
    text = text.replace("|", "\\|")
    bullet = bullet.replace("|", "\\|")
    md_lines.append(f"| {text} | {bullet} |")

with open(output_md, "w", encoding="utf-8") as f:
    f.write("\n".join(md_lines))

print(f"✅ Markdown table with fixed hierarchy saved to: {os.path.abspath(output_md)}")
