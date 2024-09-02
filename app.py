import openpyxl # leitura do arquivo excel 
from PIL import Image, ImageDraw, ImageFont # preenche os dados da planilha no certificado modelo 

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx') # abre a planilha 
sheet_alunos = workbook_alunos['Sheet1'] # seleciona a folha da planilha

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):  # looping para coletar dados 
    nome_curso = linha[0].value  # coleta o nome do curso
    nome_participante = linha[1].value # coleta o nome do participante 
    tipo_participante = linha[2].value # coleta o tipo de participante 
    data_inicio = linha[3].value # coleta a data de inicio do curso
    data_final  = linha[4].value # coleta a data final do curso
    carga_horaria = linha[5].value # coleta a carga horaria do aluno 
    data_emissao = linha [6].value # coleta a data de emissão do certificado 

# transferir de dados para o certificado 

# Definindo as fontes a ser usada 

    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90) # fonte dos nomes 
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 90) # fonte das outras especificações 
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55) # fonte das outras especificações 

    image = Image.open('./certificado_padrao.jpg') # abre o certificado pardrão para inserir as informações 
    desenhar = ImageDraw.Draw(image) # inicia a inserção das informações 

# inserindo dados no certificado 

    desenhar.text((1020,827),nome_participante,fill='black',font=fonte_nome) # indere o nome do participante
    desenhar.text((1080,940),nome_curso,fill='black',font=fonte_geral) # insere o nome do curso 
    desenhar.text((1440,1060),tipo_participante,fill='black',font=fonte_geral) # tipo de participante 
    desenhar.text((1480,1182),str(carga_horaria),fill='black',font=fonte_geral) # insere a carga horaria 
    desenhar.text((750,1770),data_inicio, fill='blue', font=fonte_data) # insere a data de inicio do curso 
    desenhar.text((750,1925),data_final, fill='blue', font=fonte_data) # insere a data final do curso 
    desenhar.text((2220,1930),data_emissao, fill='blue', font=fonte_data) # insere a data de emissão d0 certificado 
    
    
    image.save(f'./certificado/{indice} {nome_participante} .png') # salva tudo em um nove arquivo em png 