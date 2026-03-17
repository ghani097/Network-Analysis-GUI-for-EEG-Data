"""Render PlantUML diagram to PNG and SVG using PlantUML server."""
import requests
import zlib
from pathlib import Path

def encode_plantuml(text):
    """Encode PlantUML text for URL."""
    compressed = zlib.compress(text.encode('utf-8'), 9)[2:-4]
    alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_"
    result = []
    for i in range(0, len(compressed), 3):
        chunk = compressed[i:i+3]
        if len(chunk) == 3:
            b1, b2, b3 = chunk
            result.append(alphabet[b1 >> 2])
            result.append(alphabet[((b1 & 0x3) << 4) | (b2 >> 4)])
            result.append(alphabet[((b2 & 0xF) << 2) | (b3 >> 6)])
            result.append(alphabet[b3 & 0x3F])
        elif len(chunk) == 2:
            b1, b2 = chunk
            result.append(alphabet[b1 >> 2])
            result.append(alphabet[((b1 & 0x3) << 4) | (b2 >> 4)])
            result.append(alphabet[(b2 & 0xF) << 2])
        elif len(chunk) == 1:
            b1 = chunk[0]
            result.append(alphabet[b1 >> 2])
            result.append(alphabet[(b1 & 0x3) << 4])
    return ''.join(result)

# Setup paths
base_dir = Path(r"E:\GIT_HUB_MAIN\PLI-Network-Analysis-GUI")
puml_file = base_dir / "pipeline_diagram.puml"

# Read the PUML file
print(f"Reading {puml_file}...")
with open(puml_file, 'r', encoding='utf-8') as f:
    puml_content = f.read()

encoded = encode_plantuml(puml_content)

# Generate PNG
print("Generating PNG...")
png_url = f"http://www.plantuml.com/plantuml/png/{encoded}"
response = requests.get(png_url, timeout=60)
if response.status_code == 200:
    output_png = base_dir / "pipeline_diagram.png"
    with open(output_png, 'wb') as f:
        f.write(response.content)
    print(f"PNG saved: {output_png} ({len(response.content)} bytes)")

# Generate SVG (vector format for publications)
print("Generating SVG...")
svg_url = f"http://www.plantuml.com/plantuml/svg/{encoded}"
response = requests.get(svg_url, timeout=60)
if response.status_code == 200:
    output_svg = base_dir / "pipeline_diagram.svg"
    with open(output_svg, 'wb') as f:
        f.write(response.content)
    print(f"SVG saved: {output_svg} ({len(response.content)} bytes)")

print("\nDone! Files created:")
print(f"  - pipeline_diagram.png (raster, for documents)")
print(f"  - pipeline_diagram.svg (vector, for publications)")
