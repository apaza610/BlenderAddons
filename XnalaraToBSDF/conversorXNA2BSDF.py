import bpy
import csv
import os

blend_dir = bpy.path.abspath("//")
csv_path = os.path.join(blend_dir, "your_file.csv")  # Replace with your actual filename

with open(csv_path, newline='', encoding='utf-8') as csvfile:
    reader = csv.reader(csvfile)
    next(reader)  # Skip header
    for row in reader:
        mat_name = row[0].strip()
        diffuse_name   = row[1].strip() if len(row) > 1 else ""
        roughness_name = row[3].strip() if len(row) > 3 else ""
        metallic_name  = row[4].strip() if len(row) > 4 else ""
        normal_name    = row[5].strip() if len(row) > 5 else ""
        emission_name  = row[6].strip() if len(row) > 6 else ""
        alpha_name     = row[7].strip() if len(row) > 7 else ""

        mat = bpy.data.materials.get(mat_name)
        if not mat:
            print(f"Material '{mat_name}' not found.")
            continue
        if not mat.use_nodes:
            mat.use_nodes = True

        nodes = mat.node_tree.nodes
        links = mat.node_tree.links

        # Remove existing Principled BSDF nodes
        for node in nodes:
            if node.type == 'BSDF_PRINCIPLED':
                nodes.remove(node)

        # Create Principled BSDF
        bsdf_node = nodes.new(type='ShaderNodeBsdfPrincipled')
        bsdf_node.location = (400, 0)

        # Find Material Output
        output_node = next((n for n in nodes if n.type == 'OUTPUT_MATERIAL'), None)
        if output_node:
            links.new(bsdf_node.outputs['BSDF'], output_node.inputs['Surface'])

        # Helper to create Image Texture node
        def create_image_node(name, location, non_color=False):
            tex_node = nodes.new(type='ShaderNodeTexImage')
            tex_node.location = location
            if name:
                tex_path = os.path.join(blend_dir, name)
                if os.path.exists(tex_path):
                    image = bpy.data.images.load(tex_path)
                    tex_node.image = image
                    if non_color:
                        tex_node.image.colorspace_settings.name = 'Non-Color'
                else:
                    print(f"Texture file '{name}' not found for material '{mat_name}'.")
            return tex_node

        # Diffuse
        diffuse_node = create_image_node(diffuse_name, (-60, 700))
        if diffuse_node.image:
            links.new(diffuse_node.outputs['Color'], bsdf_node.inputs['Base Color'])

        # Roughness
        rough_node = create_image_node(roughness_name, (-60, 100), non_color=True)
        if rough_node.image:
            links.new(rough_node.outputs['Color'], bsdf_node.inputs['Roughness'])

        # Metallic
        metallic_node = create_image_node(metallic_name, (-60, 400), non_color=True)
        if metallic_node.image:
            links.new(metallic_node.outputs['Color'], bsdf_node.inputs['Metallic'])

        # Normal
        normal_tex_node = create_image_node(normal_name, (-350, -500), non_color=True)
        normal_map_node = nodes.new(type='ShaderNodeNormalMap')
        normal_map_node.location = (-60, -500)
        links.new(normal_tex_node.outputs['Color'], normal_map_node.inputs['Color'])
        if normal_tex_node.image:
            links.new(normal_map_node.outputs['Normal'], bsdf_node.inputs['Normal'])

        # Emission (sRGB by default)
        emission_node = create_image_node(emission_name, (-60, -800), non_color=False)
        if emission_node.image:
            links.new(emission_node.outputs['Color'], bsdf_node.inputs['Emission'])

        # Alpha
        alpha_node = create_image_node(alpha_name, (-60, -200), non_color=True)
        if alpha_node.image:
            links.new(alpha_node.outputs['Color'], bsdf_node.inputs['Alpha'])

print("All materials and texture nodes created.")