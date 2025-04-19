import os
dll_path = os.path.dirname(os.path.abspath(__file__))
os.environ["PATH"] += f";{dll_path}"

from pyadomd import Pyadomd
import pandas as pd
from auxiliar import getPortPowerBIDesktop

# Substitua pela porta correta do seu modelo Power BI Desktop
PORTA = getPortPowerBIDesktop()  # exemplo
conn_str = f"Provider=MSOLAP;Data Source=localhost:{PORTA};"

# Sua query DAX copiada do visual
dax_query = """
DEFINE
	VAR __DS0Core = 
		DISTINCT('Previsto x Realizado'[UF])

	VAR __DS0PrimaryWindowed = 
		TOPN(501, __DS0Core, 'Previsto x Realizado'[UF], 1)

EVALUATE
	__DS0PrimaryWindowed

ORDER BY
	'Previsto x Realizado'[UF]
"""

with Pyadomd(conn_str) as conn:
    with conn.cursor().execute(dax_query) as cur:
        df = pd.DataFrame(cur.fetchall(), columns=[col.name for col in cur.description])

print(df)