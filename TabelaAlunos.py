import pandas as pd
import os


class SistemaGerenciamentoAlunos:
    def __init__(self, arquivo_excel):
        self.arquivo_excel = arquivo_excel
        # Verifica se o arquivo Excel já existe
        if os.path.exists(self.arquivo_excel):
            self.df = pd.read_excel(self.arquivo_excel)
        else:
            # Cria um DataFrame vazio se o arquivo não existir
            self.df = pd.DataFrame(columns=['Nome', 'Idade', 'Curso'])

    def salvar_dados(self):
        """Salva os dados do DataFrame no arquivo Excel."""
        self.df.to_excel(self.arquivo_excel, index=False)

    def cadastrar_aluno(self, nome, idade, curso):
        """Cadastra um novo aluno no sistema."""
        novo_aluno = pd.DataFrame({'Nome': [nome], 'Idade': [idade], 'Curso': [curso]})
        self.df = pd.concat([self.df, novo_aluno], ignore_index=True)
        self.salvar_dados()
        print(f'Aluno {nome} cadastrado com sucesso!')

    def listar_alunos(self):
        """Lista todos os alunos cadastrados."""
        if self.df.empty:
            print("Nenhum aluno cadastrado.")
        else:
            print(self.df)

    def atualizar_aluno(self, index, nome=None, idade=None, curso=None):
        """Atualiza as informações de um aluno existente."""
        if index < 0 or index >= len(self.df):
            print("Índice inválido.")
            return

        if nome:
            self.df.at[index, 'Nome'] = nome
        if idade is not None:
            self.df.at[index, 'Idade'] = idade
        if curso:
            self.df.at[index, 'Curso'] = curso

        self.salvar_dados()
        print(f'Aluno na posição {index} atualizado com sucesso!')

    def excluir_aluno(self, index):
        """Exclui um aluno do sistema."""
        if index < 0 or index >= len(self.df):
            print("Índice inválido.")
            return

        aluno_removido = self.df.iloc[index]
        self.df = self.df.drop(index).reset_index(drop=True)
        self.salvar_dados()
        print(f'Aluno {aluno_removido["Nome"]} excluído com sucesso!')


def main():
    sistema = SistemaGerenciamentoAlunos('alunos.xlsx')

    while True:
        print("\nSistema de Gerenciamento de Alunos")
        print("1. Cadastrar Aluno")
        print("2. Listar Alunos")
        print("3. Atualizar Aluno")
        print("4. Excluir Aluno")
        print("5. Sair")
        opcao = input("Escolha uma opção: ")

        if opcao == '1':
            nome = input("Nome do Aluno: ")
            idade = int(input("Idade do Aluno: "))
            curso = input("Curso do Aluno: ")
            sistema.cadastrar_aluno(nome, idade, curso)

        elif opcao == '2':
            sistema.listar_alunos()

        elif opcao == '3':
            sistema.listar_alunos()  # Mostra a lista antes de atualizar
            index = int(input("Índice do Aluno para atualizar: "))
            nome = input("Novo Nome (digite 0 para não alterar): ")
            idade = input("Nova Idade (digite 0 para não alterar): ")
            curso = input("Novo Curso (digite 0 para não alterar): ")

            # Verifica se o usuário digitou '0' e não altera o valor se for o caso
            if nome == '0':
                nome = None
            if idade == '0':
                idade = None
            else:
                idade = int(idade)  # Converte para inteiro se não for '0'
            if curso == '0':
                curso = None

            sistema.atualizar_aluno(index, nome, idade, curso)

        elif opcao == '4':
            sistema.listar_alunos()  # Mostra a lista antes de excluir
            index = int(input("Índice do Aluno para excluir: "))
            sistema.excluir_aluno(index)

        elif opcao == '5':
            print("Saindo do sistema...")
            break

        else:
            print("Opção inválida. Tente novamente.")


if __name__ == "__main__":
    main()