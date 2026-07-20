import ast, sys
try:
    src = open(r"C:\Users\Pedro Lima\Documents\IMPORT_COLABORADORES\ImportarUsuariosUmuarama.py", encoding="utf-8").read()
    ast.parse(src)
    print("Sintaxe OK")
except SyntaxError as e:
    print(f"ERRO DE SINTAXE: linha {e.lineno}: {e.msg}")
    print(f"  {e.text}")
    sys.exit(1)
