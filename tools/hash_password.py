"""Vytvoří hash hesla pro Streamlit secrets (sekce [users]).

Spuštění v kořeni projektu:  python tools/hash_password.py
Heslo se zadává skrytě, nikam se neukládá a nevypisuje – vypíše se jen hash.
"""

import getpass
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import auth  # noqa: E402

password = getpass.getpass("Nové heslo: ")
if password != getpass.getpass("Heslo znovu: "):
    sys.exit("Hesla se neshodují.")
problems = auth.password_problems(password)
if problems:
    sys.exit("Heslo nevyhovuje: " + ", ".join(problems) + ".")
print(auth.hash_password(password))
