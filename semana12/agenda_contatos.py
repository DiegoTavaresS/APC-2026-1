contatos = {}

while True:

    print("\n=== AGENDA DE CONTATOS ===")
    print("1 - Adicionar contato")
    print("2 - Buscar contato")
    print("3 - Listar contatos")
    print("0 - Sair")

    opcao = input("Escolha uma opção: ")

    if opcao == "1":
        nome = input("Nome: ")
        telefone = input("Telefone: ")

        contatos[nome] = telefone
        print("Contato cadastrado!")

    elif opcao == "2":
        nome = input("Nome para busca: ")

        if nome in contatos:
            print(f"Telefone: {contatos[nome]}")
        else:
            print("Contato não encontrado.")

    elif opcao == "3":
        print("\nContatos cadastrados:")

        for nome, telefone in contatos.items():
            print(f"{nome}: {telefone}")

    elif opcao == "0":
        print("Encerrando agenda...")
        break

    else:
        print("Opção inválida!")
