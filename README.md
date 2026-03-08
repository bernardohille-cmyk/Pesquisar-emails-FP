# 🏛️ AP Contactos (Streamlit)

Aplicação **Streamlit** para carregar, filtrar, enriquecer e exportar contactos da Administração Pública portuguesa.

Permite:
- importar bases de dados Excel
- fundir dados externos (incluindo SIOE)
- pesquisar organismos
- selecionar e exportar emails
- enriquecer automaticamente dados via web

---

# 🚀 Executar

Instalar dependências:

pip install -r requirements.txt

Executar a app:

streamlit run app.py

---

# 📂 Estrutura recomendada

project/
│
├─ app.py
├─ requirements.txt
├─ README.md
│
└─ scripts/
   └─ check_diff_markers.py

---

# ⚠️ Diagnóstico rápido

Se aparecer:

SyntaxError: invalid decimal literal

e no erro existir algo como:

diff --git
index
@@

então o ficheiro Python foi contaminado com **metadados de patch/diff do Git**.

---

# 🛠️ Verificação de sintaxe

Antes de executar:

python -m py_compile app.py

Opcional:

python scripts/check_diff_markers.py app.py

---

# 📊 Funcionalidades

### Importação
- Excel base principal
- Exportação SIOE
- Bases externas

### Pesquisa
- entidade
- ministério
- categoria
- email

### Exportação
- Excel completo
- CSV filtrado
- lista de emails

---

# 📜 Licença

Uso interno / investigação.
