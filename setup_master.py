from dotenv import load_dotenv
load_dotenv()

import auth

auth.init_db()

# Cria o utilizador master com credenciais definidas
ok, err = auth.create_user('admin_lumini', 'admin123')
if ok:
    # Promove a master
    import psycopg2
    conn = auth.get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET is_master = TRUE WHERE username = 'admin_lumini'")
    conn.commit()
    cur.close()
    conn.close()
    print("Utilizador 'admin_lumini' criado como master com senha 'admin123'.")
else:
    print(f"Aviso: {err} — pode já existir. A promover a master mesmo assim...")
    conn = auth.get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET is_master = TRUE WHERE username = 'admin_lumini'")
    conn.commit()
    cur.close()
    conn.close()
    print("Utilizador 'admin_lumini' promovido a master.")
