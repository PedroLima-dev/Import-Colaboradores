import py_compile, sys
for f in ["ImportarUsuariosUmuarama.py", "gerar_importar_usuarios_uap.py"]:
    try:
        py_compile.compile(f, doraise=True)
        print(f"Sintaxe OK: {f}")
    except Exception as e:
        print(f"ERRO: {e}")
        sys.exit(1)

