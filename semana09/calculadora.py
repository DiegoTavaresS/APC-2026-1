"""
Calculadora simples
APC - Semana 09
Autor: Diego Silva
"""

def soma(a, b):
    return a + b

def subtracao(a, b):
    return a - b

def multiplicacao(a, b):
    return a * b

def divisao(a, b):
    if b == 0:
        return "Erro: divisão por zero"
    return a / b


while True:
    print("\n=== CALCULADORA ===")
    print("1 - Soma")
    print("2 - Subtração")
    print("3 - Multiplicação")
    print("4 - Divisão")
    print("0 - Sair")

    opcao = input("Escolha uma opção: ")

    if opcao == "0":
        print("Encerrando...")
        break

    a = float(input("Primeiro número: "))
    b = float(input("Segundo número: "))

    if opcao == "1":
        print("Resultado:", soma(a, b))
    elif opcao == "2":
        print("Resultado:", subtracao(a, b))
    elif opcao == "3":
        print("Resultado:", multiplicacao(a, b))
    elif opcao == "4":
        print("Resultado:", divisao(a, b))
    else:
        print("Opção inválida!")
