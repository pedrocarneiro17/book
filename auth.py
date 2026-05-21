import os
import secrets
import psycopg2
from psycopg2.extras import RealDictCursor
from werkzeug.security import generate_password_hash, check_password_hash


def get_db():
    url = os.environ.get('DATABASE_URL', '')
    # Railway expõe postgres:// mas psycopg2 precisa de postgresql://
    if url.startswith('postgres://'):
        url = url.replace('postgres://', 'postgresql://', 1)
    return psycopg2.connect(url)


def init_db():
    """Cria as tabelas e o utilizador master na primeira execução."""
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id           SERIAL PRIMARY KEY,
            username     VARCHAR(100) UNIQUE NOT NULL,
            password_hash VARCHAR(255) NOT NULL,
            is_master    BOOLEAN DEFAULT FALSE,
            is_active    BOOLEAN DEFAULT TRUE,
            created_at   TIMESTAMP DEFAULT NOW()
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS ip_logs (
            id           SERIAL PRIMARY KEY,
            user_id      INTEGER REFERENCES users(id) ON DELETE CASCADE,
            ip_address   VARCHAR(45) NOT NULL,
            user_agent   TEXT,
            accessed_at  TIMESTAMP DEFAULT NOW()
        )
    """)

    # Cria master só se não existir nenhum
    cur.execute("SELECT id FROM users WHERE is_master = TRUE LIMIT 1")
    if cur.fetchone() is None:
        password = secrets.token_urlsafe(12)
        cur.execute(
            "INSERT INTO users (username, password_hash, is_master) VALUES (%s, %s, TRUE)",
            ('master', generate_password_hash(password))
        )
        print("=" * 55)
        print("[AUTH] Utilizador master criado pela primeira vez.")
        print(f"  Username : master")
        print(f"  Password : {password}")
        print("[AUTH] Guarda esta password — não será mostrada novamente!")
        print("=" * 55)

    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def verify_login(username, password):
    """Devolve dict do utilizador ou None se credenciais inválidas."""
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute(
        "SELECT * FROM users WHERE username = %s AND is_active = TRUE",
        (username,)
    )
    user = cur.fetchone()
    cur.close()
    conn.close()
    if user and check_password_hash(user['password_hash'], password):
        return dict(user)
    return None


def get_user_by_id(user_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute(
        "SELECT id, username, is_master, is_active FROM users WHERE id = %s",
        (user_id,)
    )
    user = cur.fetchone()
    cur.close()
    conn.close()
    return dict(user) if user else None


# ---------------------------------------------------------------------------
# IP logging
# ---------------------------------------------------------------------------

def log_ip(user_id, ip_address, user_agent):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO ip_logs (user_id, ip_address, user_agent) VALUES (%s, %s, %s)",
        (user_id, ip_address, user_agent)
    )
    conn.commit()
    cur.close()
    conn.close()


def get_ip_summary():
    """
    Devolve lista de dicts com:
      username, ip_address, access_count, last_access
    Ordenado por username e depois por last_access DESC.
    """
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("""
        SELECT
            u.id         AS user_id,
            u.username,
            u.is_active,
            l.ip_address,
            COUNT(*)               AS access_count,
            MAX(l.accessed_at)     AS last_access
        FROM ip_logs l
        JOIN users u ON u.id = l.user_id
        GROUP BY u.id, u.username, u.is_active, l.ip_address
        ORDER BY u.username, last_access DESC
    """)
    rows = [dict(r) for r in cur.fetchall()]
    cur.close()
    conn.close()
    return rows


def clear_user_ips(user_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM ip_logs WHERE user_id = %s", (user_id,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Gestão de utilizadores (master only)
# ---------------------------------------------------------------------------

def get_all_users():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("""
        SELECT id, username, is_master, is_active, created_at,
               (SELECT COUNT(DISTINCT ip_address) FROM ip_logs WHERE user_id = users.id) AS ip_count
        FROM users
        ORDER BY created_at
    """)
    users = [dict(r) for r in cur.fetchall()]
    cur.close()
    conn.close()
    return users


def create_user(username, password):
    """Devolve (True, None) ou (False, mensagem_erro)."""
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO users (username, password_hash) VALUES (%s, %s)",
            (username, generate_password_hash(password))
        )
        conn.commit()
        return True, None
    except psycopg2.errors.UniqueViolation:
        conn.rollback()
        return False, "Username já existe."
    finally:
        cur.close()
        conn.close()


def delete_user(user_id, current_user_id):
    """Apaga utilizador e todos os seus logs de IP. Não apaga o próprio utilizador logado."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE id = %s AND id != %s", (user_id, current_user_id))
    conn.commit()
    cur.close()
    conn.close()


def toggle_user_active(user_id):
    """Ativa/desativa — não funciona no master."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "UPDATE users SET is_active = NOT is_active WHERE id = %s AND is_master = FALSE",
        (user_id,)
    )
    conn.commit()
    cur.close()
    conn.close()


def update_password(user_id, new_password):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "UPDATE users SET password_hash = %s WHERE id = %s",
        (generate_password_hash(new_password), user_id)
    )
    conn.commit()
    cur.close()
    conn.close()
