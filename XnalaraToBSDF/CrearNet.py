import bpy

# --- Get or create an active material ---
obj = bpy.context.object
if not obj:
    raise Exception("No active object selected!")

mat = obj.active_material
if not mat:
    mat = bpy.data.materials.new(name="XPS_Converted")
    obj.active_material = mat

mat.use_nodes = True
nodes = mat.node_tree.nodes
links = mat.node_tree.links

# --- Find or create the Material Output node ---
output = next((n for n in nodes if n.type == "OUTPUT_MATERIAL"), None)
if not output:
    output = nodes.new("ShaderNodeOutputMaterial")
output.location = (800, 0)

# --- Find or create Principled BSDF node ---
bsdf = next((n for n in nodes if n.type == "BSDF_PRINCIPLED"), None)
if not bsdf:
    bsdf = nodes.new("ShaderNodeBsdfPrincipled")
    bsdf.location = (400, 0)

# Link BSDF → Output if not already linked
if not any(l.to_node == output and l.from_node == bsdf for l in links):
    links.new(bsdf.outputs["BSDF"], output.inputs["Surface"])

# --- Helper: find or create image texture node by label/name ---
def get_or_create_image_node(label, x, y):
    node = next((n for n in nodes if n.label == label or n.name == label), None)
    if not node:
        node = nodes.new("ShaderNodeTexImage")
        node.label = label
        node.name = label
        node.location = (x, y)
        node.interpolation = "Linear"
    return node

# --- Ensure required image nodes exist ---
basecolor = get_or_create_image_node("BaseColor", -800, 200)
metallic  = get_or_create_image_node("Metallic",  -800,  50)
roughness = get_or_create_image_node("Roughness",-800, -100)
alpha     = get_or_create_image_node("Alpha",    -800, -250)
normalimg = get_or_create_image_node("NormalGL", -800, -450)

# --- Reuse Diffuse texture for BaseColor ---
diffuse_node = next((n for n in nodes if n.label == "Diffuse" or n.name == "Diffuse"), None)
if diffuse_node and hasattr(diffuse_node, "image") and diffuse_node.image:
    basecolor.image = diffuse_node.image
    print(f"🎨 Reused Diffuse texture for BaseColor: {diffuse_node.image.name}")

# --- Find or create Normal Map node ---
normalmap = next((n for n in nodes if n.type == "NORMAL_MAP"), None)
if not normalmap:
    normalmap = nodes.new("ShaderNodeNormalMap")
    normalmap.location = (-400, -450)
    normalmap.space = "TANGENT"

# --- Safe link helper (avoids duplicates) ---
def safe_link(output_socket, input_socket):
    for l in links:
        if l.to_socket == input_socket:
            return  # already linked
    links.new(output_socket, input_socket)

# --- Connect nodes (only if not connected yet) ---
safe_link(basecolor.outputs["Color"], bsdf.inputs["Base Color"])
safe_link(metallic.outputs["Color"], bsdf.inputs["Metallic"])
safe_link(roughness.outputs["Color"], bsdf.inputs["Roughness"])
safe_link(alpha.outputs["Color"], bsdf.inputs["Alpha"])   # ✅ Correct Alpha link
safe_link(normalimg.outputs["Color"], normalmap.inputs["Color"])
safe_link(normalmap.outputs["Normal"], bsdf.inputs["Normal"])

print("✅ Node network ensured: reused existing nodes and Diffuse texture where available.")
