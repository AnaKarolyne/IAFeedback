import speech_recognition as sr
import pyttsx3
import os
import openpyxl
from openpyxl import Workbook
import datetime

recognizer = sr.Recognizer()
engine = pyttsx3.init()

def maquina_fala(fala):
    print(fala)
    engine.say(fala)
    engine.runAndWait()

def unknown_value():
    erro = "Não foi possível entender a sua fala. Tente novamente."
    maquina_fala(erro)
    return

def request_error():
    erro = "Ocorreu um erro na API de reconhecimento de fala: {e}"
    maquina_fala(erro)
    return

def mostrar_feedbacks():
    # Nome do arquivo Excel a ser lido
    arquivo_excel = "feedbacks.xlsx"

    try:
        workbook = openpyxl.load_workbook(arquivo_excel)
    except FileNotFoundError:
        erro = "O arquivo Excel ainda não existe ou foi excluído."
        maquina_fala(erro)
        return

    # Abre a planilha padrão (Sheet1)
    sheet = workbook.active

    # Verifica se a planilha não está vazia
    if sheet.max_row <= 1:
        erro = "Não há registros de feedback no arquivo Excel."
        maquina_fala(erro)
        return

    # Itera pelas linhas do arquivo Excel e exibe os registros
    registros = "Registros de Feedback:"
    maquina_fala(registros)

    for row in sheet.iter_rows(min_row=2, values_only=True):
        #Única string com todos os valores da linha
        linha_completa = " | ".join(map(str, row))
        print(linha_completa)

    workbook.close()

def coletar_usabilidade():
    usabilidade1 = input("Como você avaliaria a facilidade de uso do software em uma escala de 1 a 10? ")
    usabilidade2 = input("Quais aspectos específicos da interface do usuário você acha que podem ser melhorados? ")
    usabilidade3 = input("Você teve algum problema ao navegar pelo software? Se sim, por favor, descreva. ")
    return usabilidade1, usabilidade2, usabilidade3

def coletar_geral():
    geral1 = input("Em uma escala de 1 a 10, quão satisfeito você está com o software? ")
    geral2 = input("Você recomendaria este software a outros profissionais? Por quê? ")
    geral3 = input("Existe algum feedback adicional que você gostaria de compartilhar sobre sua experiência geral com o software? ")
    return geral1, geral2, geral3

def coletar_atualizacoes():
    atualizacoes1 = input("Como você se sente em relação às atualizações frequentes do software? ")
    atualizacoes2 = input("Há alguma atualização recente que você achou particularmente útil ou problemática? ")
    atualizacoes3 = input("Que tipo de melhorias você gostaria de ver nas próximas versões do software? ")
    return atualizacoes1, atualizacoes2, atualizacoes3

def coletar_seguranca():
    seguranca1 = input("Você se sente seguro ao usar este software em relação à segurança de seus dados? ")
    seguranca2 = input("Existe alguma preocupação específica de segurança ou privacidade que você gostaria de mencionar? ")
    seguranca3 = input("Você percebeu alguma vulnerabilidade de segurança enquanto usava o software? ")
    return seguranca1, seguranca2, seguranca3

def coletar_suporte():
    suporte1 = input("Como você avalia a qualidade do suporte técnico fornecido pela empresa? ")
    suporte2 = input("Você teve uma boa experiência com o atendimento ao cliente? Pode compartilhar detalhes? ")
    suporte3 = input("Existe algum incidente de suporte técnico que você gostaria de destacar? ")
    return suporte1, suporte2, suporte3


def coletar_recursos():
    recursos1 = input("Quais recursos ou funcionalidades do software você considera mais úteis? ")
    recursos2 = input("Há alguma funcionalidade que você gostaria que o software tivesse, mas não tem atualmente? ")
    recursos3 = input("Você teve dificuldades em encontrar alguma função específica no software? ")
    return recursos1, recursos2, recursos3

def coletar_desempenho():
    desempenho1 = input("O software atendeu às suas expectativas em termos de desempenho e velocidade? ")
    desempenho2 = input("Você notou alguma lentidão ou travamento ao usar o software? ")
    desempenho3 = input("Em que situações o software parece mais lento para você? ")
    return desempenho1, desempenho2, desempenho3

# Função para coletar feedback e armazená-lo no arquivo Excel
def coletar_feedback():

    nome_cliente = input("Qual é o seu nome? ")
    nome_produto = input(f"Qual é o nome do produto que você vai avaliar {nome_cliente}?")

    try:
        # Chamando a função coletar_usabilidade
        usabilidade1, usabilidade2, usabilidade3 = coletar_usabilidade()

        desempenho1, desempenho2, desempenho3 = coletar_desempenho()
        recursos1, recursos2, recursos3 = coletar_recursos()
        suporte1, suporte2, suporte3 = coletar_suporte()
        seguranca1, seguranca2, seguranca3 = coletar_seguranca()
        atualizacoes1, atualizacoes2, atualizacoes3 = coletar_atualizacoes()
        geral1, geral2, geral3 = coletar_geral()

        salvar_feedback(nome_cliente, nome_produto, usabilidade1, usabilidade2, usabilidade3, desempenho1, desempenho2, desempenho3, recursos1, recursos2, recursos3, suporte1, suporte2, suporte3, seguranca1, seguranca2, seguranca3, atualizacoes1, atualizacoes2, atualizacoes3, geral1, geral2, geral3)

    except sr.UnknownValueError:
        unknown_value()
        coletar_feedback()

    except sr.RequestError as e:
        request_error()
        coletar_feedback()

def salvar_feedback(nome_cliente, nome_produto, usabilidade1, usabilidade2, usabilidade3, desempenho1, desempenho2, desempenho3, recursos1, recursos2, recursos3, suporte1, suporte2, suporte3, seguranca1, seguranca2, seguranca3, atualizacoes1, atualizacoes2, atualizacoes3, geral1, geral2, geral3):

    # Nome do arquivo Excel que você deseja criar ou verificar
    nome_arquivo = "feedbacks.xlsx"

    # Verifica se o arquivo já existe no diretório atual
    if os.path.isfile(nome_arquivo):
        print(f"O arquivo '{nome_arquivo}' já existe.")

    else:
        try:
            # Crie um novo arquivo Excel (workbook)
            workbook = openpyxl.Workbook()

            # Adicione uma planilha (worksheet) padrão
            sheet = workbook.active

            sheet["A1"] = "Data e Hora"
            sheet["B1"] = "Nome do Cliente"
            sheet["C1"] = "Nome do Produto Avaliado"
            sheet["D1"] = "Usabilidade1"
            sheet["E1"] = "Usabilidade2"
            sheet["F1"] = "Usabilidade3"
            sheet["G1"] = "Desempenho1"
            sheet["H1"] = "Desempenho2"
            sheet["I1"] = "Desempenho3"
            sheet["J1"] = "Recursos1"
            sheet["K1"] = "Recursos2"
            sheet["L1"] = "Recursos3"
            sheet["M1"] = "Suporte1"
            sheet["N1"] = "Suporte2"
            sheet["O1"] = "Suporte3"
            sheet["P1"] = "Segurança1"
            sheet["Q1"] = "Segurança2"
            sheet["R1"] = "Segurança3"
            sheet["S1"] = "Atualizações1"
            sheet["T1"] = "Atualizações2"
            sheet["U1"] = "Atualizações3"
            sheet["V1"] = "Geral1"
            sheet["W1"] = "Geral2"
            sheet["X1"] = "Geral3"

            # Salve o arquivo Excel com um nome específico
            workbook.save("feedbacks.xlsx")

            # Feche o arquivo (não é estritamente necessário, mas é uma boa prática)
            workbook.close()

            print(f"O arquivo feedbacks.xlsx foi criado com sucesso.")

        except FileNotFoundError:
            workbook = Workbook()

    # Abre o arquivo Excel existente
    workbook = openpyxl.load_workbook(nome_arquivo)

    # Selecione a planilha existente ou crie uma nova, se necessário
    if "Sheet" in workbook.sheetnames:
        sheet = workbook["Sheet"]
    else:
        sheet = workbook.create_sheet("Sheet")

    # Obtém a data e hora atual
    data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Adiciona o feedback à planilha
    nova_linha = [data_hora, nome_cliente, nome_produto, usabilidade1, usabilidade2, usabilidade3, desempenho1, desempenho2, desempenho3, recursos1, recursos2, recursos3, suporte1, suporte2, suporte3, seguranca1, seguranca2, seguranca3, atualizacoes1, atualizacoes2, atualizacoes3, geral1, geral2, geral3]
    sheet.append(nova_linha)

    # Salva as alterações no arquivo Excel
    workbook.save(nome_arquivo)

    # Feche o arquivo (não é estritamente necessário, mas é uma boa prática)
    workbook.close()

    sucesso = "Feedback armazenado com sucesso!"
    maquina_fala(sucesso)
    principal()

# Função principal - Íncio da trilha
def trilha():
    comeco = "Bem vindo a nossa trilha de Feedback."
    coletar_feedback()

def funcoes():
    while True:
        with sr.Microphone() as source:
            comando = input("Eu fui projetada especialmente para te ouvir, entender quais são as suas críticas em relação ao nosso produto e descobrir, com a sua ajuda, como podemos fornecer um produto e atendimento melhor para você. Você gostaria de dar um feedback ou visualizar os feedbacks coletados?")

        try:
            if "dar" in comando or "novo" in comando:
                trilha()
            elif "visualizar" in comando or "Visualizar" in comando or "ver" in comando or "analisar" in comando:
                mostrar_feedbacks()
            elif "não" in comando or "Não" in comando or "nenhuma" in comando:
                negativa = "Tudo bem, espero te ver numa próxima vez e que você tenha uma experiência melhor com o nosso produto."
                principal()
            else:
                erro = "Não entendi o comando. Tente novamente."
                funcoes()
        except sr.UnknownValueError:
            unknown_value()

        except sr.RequestError as e:
            request_error()

def principal():
    while True:
        with sr.Microphone() as source:
            comando = input("Olá, eu sou a IA VozDoCliente, uma inteligência artificial desenvolvida especialmente para te ouvir. Vamos começar? ")

            try:
                if "sim" in comando or "começar" in comando or "Sim" in comando:
                    mestre = "Fico feliz em ajudar e a colaborar para o desenvolvimento de produtos mais assertivos."
                    funcoes()
                else:
                    mestre = "Não entendi... Talvez eu não seja a IA que você procure."
                    principal()
            except sr.UnknownValueError:
                unknown_value()

            except sr.RequestError as e:
                request_error()

principal()