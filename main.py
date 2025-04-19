import os
dll_path = os.path.dirname(os.path.abspath(__file__))
os.environ["PATH"] += f";{dll_path}"

from pyadomd import Pyadomd

# Substitua pelos seus dados:
workspace_url = "powerbi://api.powerbi.com/v1.0/myorg/Opera%C3%A7%C3%B5es%20Cobre"
dataset_name = "Dash Acompanhamento Cobre"

conn_str = f"Data Source={workspace_url};Initial Catalog={dataset_name}"

query_dax = """
EVALUATE
TOPN(5, 'Previsto x Realizado')
"""

with Pyadomd(conn_str) as conn:
    with conn.cursor().execute(query_dax) as cur:
        columns = [col.name for col in cur.description]
        rows = [dict(zip(columns, row)) for row in cur.fetchall()]
        for row in rows:
            print(row)