import sqlite3 as sql
import textwrap as tw
import datetime as dt
from abc import ABC, abstractmethod

#DATABASE

def conectar_banco():

    try:
        connect = sql.connect("banco.db")
        connect.execute("PRAGMA foreign_keys = ON;")
        return connect
    except sql.Error as e:
        print(f"Erro ao conectar o banco: {e}")
        return None

def criar_tabela(connect):
    cursor = connect.cursor()
    clientes = """ 
        CREATE TABLE IF NOT EXISTS clientes (
            cliente INTEGER PRIMARY KEY AUTOINCREMENT,
            nome VARCHAR(100) NOT NULL,
            cpf VARCHAR(11) NOT NULL UNIQUE
        );"""
    
    contas = """
        CREATE TABLE IF NOT EXISTS contas(
            conta INTEGER PRIMARY KEY AUTOINCREMENT,
            numero INTEGER NOT NULL UNIQUE,
            agencia VARCHAR(4) NOT NULL DEFAULT'0001',
            saldo NUMERIC NOT NULL DEFAULT 0.0,
            tipo_conta VARCHAR(30) NOT NULL CHECK(tipo_conta IN ('conta corrente', 'conta poupanca')),
            limite_saques_diarios INTEGER DEFAULT 3,
            limite_valor_saque NUMERIC DEFAULT 1000.0,
            cliente_conta INTEGER NOT NULL,
            FOREIGN KEY (cliente_conta) REFERENCES clientes(cliente) 
        );"""
    
    transacoes = """ 
        CREATE TABLE IF NOT EXISTS transacoes (
            transacao INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo VARCHAR(10) NOT NULL CHECK(tipo IN ('Deposito', 'Saque')),
            valor NUMERIC NOT NULL,
            data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            conta_transacao INTEGER NOT NULL,
            FOREIGN KEY (conta_transacao) REFERENCES contas(conta)
        );"""

    cursor.execute(clientes)
    cursor.execute(contas)
    cursor.execute(transacoes)
    connect.commit()
    print("Tabelas criadas com sucesso!")
    cursor.close()


#MENU

def menu():
    menu = """\n
    ================ MENU ================
    [1]\tDepositar
    [2]\tSacar
    [3]\tExtrato
    [4]\tNova conta
    [5]\tListar contas
    [6]\tSelecionar conta
    [7]\tVer conta ativa
    [8]\tNovo Cliente
    [0]\tSair
    => """
    return input(tw.dedent(menu))

#CLASSES

class Cliente:
    def __init__(self, nome, cpf):
        self.nome = nome
        self._cpf = cpf
        self.contas = []
        self.cliente_id = None

    @property
    def cpf(self):
        return self._cpf
        
    @property
    def cpf_formatado(self):
        cpf = self._cpf
        return f"{cpf[0:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}"

    def adicionar_conta(self, conta):
        self.contas.append(conta)
        print(f"Nova Conta ({conta.tipo_conta()}) adicionada para o cliente {self.nome}")


class Conta(ABC):

    AGENCIA = "0001"
    
    def __init__(self, numero, cliente):
        self._saldo = 0.0
        self._numero = numero
        self._cliente = cliente
        self._limite_saque_diario = 3
        self._saques_realizados = 0
        self._data_saque = None
        self.conta_id = None
        print(f"Conta {self.AGENCIA}-{self.numero} criada com sucesso!")

    @property
    def saldo(self):
        return self._saldo
    
    @property
    def numero(self):
        return self._numero
    
    @property
    def cliente(self):
        return self._cliente
    
    def _saque_reset(self):
        hoje = dt.date.today()
        if self._data_saque != hoje:
            self._saques_realizados = 0
            self._data_saque = hoje

    def tratamento_erros_depositar(self):
        while True:
            try:
                valor = float(input("Qual o valor do deposito: "))
                if valor > 0:
                    return valor
                else:
                    print("O valor do deposito deve ser maior que zero.")

            except ValueError:
                print("Valor invalido. Por favor, insira um numero valido.")

    def tratamento_erros_sacar(self):
        while True:
            try:
                valor = float(input("qual o valor do saque: "))
                if valor > 0:
                    if self.saldo >= valor:
                        return valor
                    else:
                        return None
                else:
                    print ("O valor do saque deve ser maior que 0")

            except ValueError:
                print("Valor invalido. Por favor, insira um numero valido.")

    def depositar(self, connect):
        valor = self.tratamento_erros_depositar()

        connect.execute("UPDATE contas SET saldo = saldo + ? Where conta = ?", (valor, self.conta_id))

        connect.execute("INSERT INTO transacoes (tipo, valor, conta_transacao) VALUES (?, ?, ?)", ("Deposito", valor, self.conta_id))

        connect.commit()

        self._saldo += valor
        print("\n Deposito realizado com sucesso!")
    
    def sacar(self, connect):

        self._saque_reset()

        if self._saques_realizados >= self._limite_saque_diario:
            print(f"Falha, limite de {self._limite_saque_diario} saques diarios atingidos")
            return False

        valor = self.tratamento_erros_sacar()

        if valor:

            connect.execute("UPDATE contas SET saldo = saldo - ? Where conta = ?", (valor, self.conta_id))

            connect.execute("INSERT INTO transacoes (tipo, valor, conta_transacao) VALUES (?, ?, ?)", ("Saque", valor, self.conta_id))

            connect.commit()

            self._saldo -= valor
            self._saques_realizados += 1
            print("\n Saque realizado com sucesso!")
        else:
            print("\nSaldo insuficiente para o saque.")

    def exibir_extrato(self, connect):

        cursor = connect.cursor()

        print(f"\nEXTRATO ({self.tipo_conta()})")
        print(f"\nCliente: {self._cliente.nome}")
        print(f"\nConta: {self.AGENCIA}-{self.numero}")

        cursor.execute("SELECT tipo, valor, data_hora FROM transacoes WHERE conta_transacao = ?", (self.conta_id,))

        transacoes = cursor.fetchall()

        if not transacoes:
            print("\nNão foi feito nenhuma movimentação na conta\n")
        else:
            for i, j, k in transacoes:

                data_obj = dt.datetime.strptime(k, '%Y-%m-%d %H:%M:%S')
                data_formatada = data_obj.strftime("%d/%m/%Y %H:%M:%S")
                cor = '\033[92m' if i == "Deposito" else '\033[91m'
                sinal = '+' if i == "Deposito" else '-'
                print(f"{data_formatada} - {i:8} {cor}\t {sinal} R$ {j:>10.2f}\033[0m")

        print(f"Saldo atual é de {self.saldo:.2f}")

    @abstractmethod
    def tipo_conta(self):
        pass

class ContaCorrente(Conta):
    def __init__(self, numero, cliente):
        super().__init__(numero, cliente)
        self.limite_valor_saque = 1000.0

    def sacar(self, connect):

        self._saque_reset()

        if self._saques_realizados >= self._limite_saque_diario:
            print(f"Falha, limite de {self._limite_saque_diario} saques diarios atingidos")
            return False

        valor = self.tratamento_erros_sacar()

        if valor:

            if valor > self.limite_valor_saque:
                print(f"Falha! O valor (R${valor:.2f}) passou do limite por saque R$ 1000.")
                return False
            
            connect.execute("UPDATE contas SET saldo = saldo - ? Where conta = ?", (valor, self.conta_id))

            connect.execute("INSERT INTO transacoes (tipo, valor, conta_transacao) VALUES (?, ?, ?)", ("Saque", valor, self.conta_id))

            connect.commit()

            self._saldo -= valor
            self._saques_realizados += 1
            self.limite_valor_saque -= valor
            print("\n Saque realizado com sucesso!")
        else:
            print("\nSaldo insuficiente para o saque. ")

    def tipo_conta(self):
        return "Conta Corrente"

class ContaPoupanca(Conta):
    def __init__(self, numero, cliente):
        super().__init__(numero, cliente)
        
    def tipo_conta(self):
        return "Conta Poupanca"

def filtrar_cliente(cpf, clientes):
    for i in clientes:
        if i.cpf == cpf:
            return i
    return None
        
def cpf_limpo(cpf):
    cpf = "".join(filter(str.isdigit, cpf))

    if len(cpf) == 11:
        return cpf
    else:
        print("CPF invalido, deve ter 11 números")
        return None

def criar_cliente(connect, clientes):
    cpf_entrada = input("\nDigite o seu CPF: ")
    cpf = cpf_limpo(cpf_entrada)

    if not cpf:
        return None

    cliente_existe = filtrar_cliente(cpf, clientes)

    if cliente_existe:
        print(f"Já existe um cliente com o cpf: {cliente_existe.cpf_formatado}") 
        return None
    
    nome = input ("\nDigite o seu nome: ")

    cursor = connect.cursor()

    cursor.execute("INSERT INTO clientes (nome, cpf) VALUES (?, ?)", (nome, cpf))

    connect.commit()

    cliente_id = cursor.lastrowid

    novo_cliente = Cliente(nome, cpf)
    novo_cliente.cliente_id = cliente_id
    clientes.append(novo_cliente)
    print(f"Cliente: {nome} (CPF: {novo_cliente.cpf_formatado}) criado com sucesso")
    return novo_cliente
    
def criar_conta(connect, clientes, contas):
    cpf_entrada = input("Digite o cpf do cliente para vincular a conta: ")
    cpf = cpf_limpo(cpf_entrada)

    if not cpf:
        return None

    cliente = filtrar_cliente(cpf, clientes)

    if not cliente:
        print("Cliente nao encontrado, use a opção 8.")
        return None

    while True:
        tipo = input("Qual o tipo da conta? [1] Corrente | [2] Poupança: ")
        if tipo in ("1", "2"):
            break
        print("Opcao invalida")

    numero_conta = len(contas) + 1
    nova_conta = None
    tipo_str = ""

    if tipo == "1":
        nova_conta = ContaCorrente(numero_conta, cliente)  
        tipo_str = nova_conta.tipo_conta().lower()  
    elif tipo == "2":
        nova_conta = ContaPoupanca(numero_conta, cliente)
        tipo_str = nova_conta.tipo_conta().lower() 

    cursor = connect.cursor()

    cursor.execute("INSERT INTO contas (numero, tipo_conta, cliente_conta, agencia, saldo, limite_saques_diarios, limite_valor_saque) VALUES (?, ?, ?, ?, ?, ?, ?)", (numero_conta, tipo_str, cliente.cliente_id, nova_conta.AGENCIA, nova_conta.saldo, nova_conta._limite_saque_diario, getattr(nova_conta, 'limite_valor_saque', 1000.0)))

    connect.commit()

    conta_id = cursor.lastrowid

    nova_conta.conta_id = conta_id 
        
    contas.append(nova_conta)
    cliente.adicionar_conta(nova_conta)
    return nova_conta

def listar_contas(contas):
    print("\n Todas as contas:")
    if not contas:
        print("Nenhuma conta cadastrada, use a opção 4.")
    
    for i in contas:
        print(f"Tipo: {i.tipo_conta()}")
        print(f"Cliente: {i.cliente.nome}")
        print(f"Agencia: {i.AGENCIA} | Conta: {i.numero}")
        print(f"Saldo: R$ {i.saldo:.2f}")
        
def selecionar_contas(clientes):
    cpf_entrada = input("Digite o cpf do cliente: ")
    cpf = cpf_limpo(cpf_entrada)

    if not cpf:
        return None

    cliente = filtrar_cliente(cpf, clientes)
    
    if not cliente:
        print("Cliente nao encontrado, use a opção 8.")
        return None
        
    if not cliente.contas:
        print(f"Cliente {cliente.nome} não possui contas")
        return None
    
    print(f"Contas do Cliente {cliente.nome}:")
    
    for i, conta in enumerate(cliente.contas):
        print(f"  [{i+1}] {conta.tipo_conta()} - Nº {conta.numero} | Saldo R$ {conta.saldo:.2f}")
        
    while True:
        try:
            escolha = int(input("Qual conta deseja selecionar? "))
            if 1 <= escolha <= len(cliente.contas):
                conta_selecionada = cliente.contas[escolha - 1]
                print(f"Conta {conta_selecionada.numero} selecionada") 
                return conta_selecionada
            else:
                print("Numero invalido")
        except ValueError:
            print("Digite um numero")   

def recuperar_dados_banco(connect):
    cursor = connect.cursor()
    lista_clientes = []
    lista_contas = []
    mapa_clientes = {} 

    print("Carregando dados...")

    # 1. Recuperar Clientes
    cursor.execute("SELECT cliente, nome, cpf FROM clientes")
    for i in cursor.fetchall():
        id_db, nome, cpf = i
        novo_cliente = Cliente(nome, cpf)
        novo_cliente.cliente_id = id_db
        lista_clientes.append(novo_cliente)
        mapa_clientes[id_db] = novo_cliente 

    # 2. Recuperar Contas
    cursor.execute("SELECT conta, numero, saldo, tipo_conta, cliente_conta FROM contas")
    for i in cursor.fetchall():
        id_conta, numero, saldo, tipo, id_dono = i
        dono = mapa_clientes.get(id_dono)
        
        if dono:
            if tipo == 'conta corrente': 
                nova_conta = ContaCorrente(numero, dono)
            else:
                nova_conta = ContaPoupanca(numero, dono)
            
            nova_conta.conta_id = id_conta
            nova_conta._saldo = float(saldo)
            
            lista_contas.append(nova_conta)
            dono.adicionar_conta(nova_conta)

    return lista_clientes, lista_contas


#MAIN

def main():

    connect = conectar_banco()
  
    criar_tabela(connect)

    clientes = []
    contas = []

    clientes, contas = recuperar_dados_banco(connect)

    conta_ativa = None

    while True:
        opcao = menu()

        if opcao == "1":
            if not conta_ativa:
                print("\nNenhuma conta selecionada, use a opção 6.")
                continue

            conta_ativa.depositar(connect)

        elif opcao == "2":
            if not conta_ativa:
                print("\nNenhuma conta selecionada, use a opção 6.")
                continue

            conta_ativa.sacar(connect)
            
        elif opcao == "3":
            if not conta_ativa:
                print("\nNenhuma conta selecionada, use a opção 6.")
                continue

            conta_ativa.exibir_extrato(connect)

        elif opcao == "4": 
            nova_conta = criar_conta(connect, clientes, contas)
            if nova_conta:
                conta_ativa = nova_conta
        
        elif opcao == "5":
            listar_contas(contas)

        elif opcao == "6":
            conta_selecionada = selecionar_contas(clientes)
            if conta_selecionada:
                conta_ativa = conta_selecionada

        elif opcao == "7":
              if not conta_ativa:
                print("\nNenhuma conta selecionada, use a opção 6.")
                continue
              else:
                print(f"Cliente: {conta_ativa.cliente.nome}")
                print(f"Tipo: {conta_ativa.tipo_conta()}")
                print(f"Agencia: {conta_ativa.AGENCIA} | Conta: {conta_ativa.numero}")
                print(f"Saldo: R$ {conta_ativa.saldo:.2f}")

        elif opcao == "8":
            criar_cliente(connect, clientes)

        elif opcao == "0": 
            print("Encerrando o sistema... Até logo!")
            connect.close()
            break

        else:
            print("\nOpção invalida")
    
main()