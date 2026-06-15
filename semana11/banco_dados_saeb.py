"""
Banco de Dados Simples (SAEB)
APC - Semana 11
Autor: Diego Tavares Silva
"""

escolas = [
    {"nome": "Escola A", "media": 250},
    {"nome": "Escola B", "media": 275},
    {"nome": "Escola C", "media": 230}
]

print("=== ESCOLAS CADASTRADAS ===")

for escola in escolas:
    print(f"{escola['nome']} - Média SAEB: {escola['media']}")

busca = input("\nDigite o nome da escola: ")

encontrada = False

for escola in escolas:
    if escola["nome"].lower() == busca.lower():
        print("\nEscola encontrada:")
        print(f"Nome: {escola['nome']}")
        print(f"Média SAEB: {escola['media']}")
        encontrada = True

if not encontrada:
    print("Escola não encontrada.")
