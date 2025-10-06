import bpy
import csv

# Output path
output_path = bpy.path.abspath("//materials_group_inputs.csv")

# Helper to get texture name from a Group node input
def get_texture_name_from_input(mat, input_name):
    if not mat.use_nodes:
        return ""
    for node in mat.node_tree.nodes:
        if node.type == 'GROUP' and input_name in node.inputs:
            socket = node.inputs[input_name]
            if socket.is_linked:
                from_node = socket.links[0].from_node
                if from_node.type == 'TEX_IMAGE' and from_node.image:
                    return from_node.image.name
    return ""

# Get materials used in scene
used_materials = set()
for obj in bpy.context.scene.objects:
    for slot in obj.material_slots:
        if slot.material:
            used_materials.add(slot.material)

# Prepare data
header = ["Material","Diffuse","Lightmap","Specular","Emission","Bump Map","Environment","Alpha"]
rows = [header]

for mat in used_materials:
    row = [
        mat.name,
        get_texture_name_from_input(mat, "Diffuse"),
        get_texture_name_from_input(mat, "Lightmap"),
        get_texture_name_from_input(mat, "Specular"),
        get_texture_name_from_input(mat, "Emission"),
        get_texture_name_from_input(mat, "Bump Map"),
        get_texture_name_from_input(mat, "Environment"),
        get_texture_name_from_input(mat, "Alpha")
    ]
    rows.append(row)

# Write CSV
with open(output_path, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerows(rows)

print(f"CSV saved to: {output_path}")
