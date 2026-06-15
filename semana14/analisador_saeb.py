escolas = [
    {"nome": "Escola A", "media": 250},
    {"nome": "Escola B", "media": 310},
    {"nome": "Escola C", "media": 280},
    {"nome": "Escola D", "media": 340}
]

# Ordenação (Bubble Sort)

for i in range(len(escolas)):
    for j in range(len(escolas) - 1):

        if escolas[j]["media"] < escolas[j + 1]["media"]:

            temp = escolas[j]
            escolas[j] = escolas[j + 1]
            escolas[j + 1] = temp

print("=== RANKING SAEB ===\n")

posicao = 1

for escola in escolas:
    print(
        f"{posicao}º lugar - "
        f"{escola['nome']} "
        f"(Média: {escola['media']})"
    )
    posicao += 1

# Média geral

soma = 0

for escola in escolas:
    soma += escola["media"]

media_geral = soma / len(escolas)

print("\nMédia Geral:", round(media_geral, 2))
