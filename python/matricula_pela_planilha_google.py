# -*- coding: utf-8 -*-
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from pyasn1.type.univ import Null
import requests
import csv
import collections
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import mysql.connector
import csv
from datetime import datetime
import os
import mimetypes
from uuid import uuid4
import time
import hashlib

mydb = mysql.connector.connect(
  #host="wiquadro-mysql-production.cmq3girrwefu.sa-east-1.rds.amazonaws.com", 
  host='172.18.0.3',
  user="wiquadro",
  password='admin',
  #password="w1qu4dr0##",  
  database="wiquadro"
)

cursor = mydb.cursor(buffered=True)

data_e_hora_atuais = datetime.now()
data_e_hora_em_texto = data_e_hora_atuais.strftime('%Y-%m-%d %H:%M:%S')

#cabeçalho da planilha
#['Carimbo de data/hora', 'Nome completo:', 'Situação acadêmica:', 'Telefone:', 'Email:', 'Cidade onde mora:', 'Como irá fazer o curso: ', 'Telecentros disponíveis:', 'CPF', 'CNPJ', 'Data de nascimento', 'Rua/Avenida:', 'Número:', 'CEP:', 'Complemento:', 'Bairro:']

disciplinas ={
    "Empreendedorismo":"01-Empreendedorismo 2021",
    "Mundo dos Micro e Pequenos Negócios":"02-Mundo dos Micro e Pequenos Negócios 2021",
    "Estratégias Digitais":"03-Estratégias Digitais p/ Pequenos Empreendedores",
    "Plano de Negócio":"04-Plano de Negócio 2021"
}

def validate_cpf(cpf):
    ''' Expects a numeric-only CPF string. '''
    TAMANHO_CPF = 11
    cpf = cpf.strip()
    if len(cpf) <  TAMANHO_CPF:
        return False  

    if cpf in (c * TAMANHO_CPF for c in "1234567890"):
        return False

    cpf_reverso = cpf[::-1]
    for i in range(2, 0, -1):
        cpf_enumerado = enumerate(cpf_reverso[i:], start=2)
        dv_calculado = sum(map(lambda x: int(x[1]) * x[0], cpf_enumerado)) * 10 % 11
        if cpf_reverso[i - 1:i] != str(dv_calculado % 10):
            return False

    return True

  
    '''    
    if cpf in [s * 11 for s in [str(n) for n in range(10)]]:
        return False
    
    calc = lambda t: int(t[1]) * (t[0] + 2)
    d1 = (sum(map(calc, enumerate(reversed(cpf[:-2])))) * 10) % 11
    d2 = (sum(map(calc, enumerate(reversed(cpf[:-1])))) * 10) % 11
    return str(d1) == cpf[-2] and str(d2) == cpf[-1]
    '''

def envia_email_aluno_cadastrado(email,message,subject):
    msg = MIMEMultipart()
    server = smtplib.SMTP_SSL('smtp.yandex.com.tr', 465)
    msg['From'] = "alex@wilivro.com.br"
    password = "123654bb"
    html = "<!DOCTYPE html>"
    html = html+"<!-- Last Published: Thu Jul 06 2017 19:06:15 GMT+0000 (UTC) -->"
    html = html+"<html>"
    html = html+"<head>"
    html = html+"<meta charset=\"utf-8\" />"
    html = html+"<title>"
    html = html+"#EmpreendeAlagoas"
    html = html+"</title>"
    html = html+"<meta content=\"width=device-width, initial-scale=1\" name=\"viewport\" />"
    html = html+"<meta name=\"generator\" />"
    html = html+"<meta name=\"facebook-domain-verification\" content=\"g0bjnhx5g2q3is02qpvqlk1cncrixk\" />"
    html = html+"<link rel=\"shortcut icon\" href=\"favicon.ico\" type=\"image/x-icon\" />"
    html = html+"<link href=\"https://empreendealagoas.com.br/src/estilo.css\" rel=\"stylesheet\" type=\"text/css\" />"
    html = html+"<script src=\"https://empreendealagoas.com.br/src/js/fonts.js\">"
    html = html+"<script type=\"text/javascript\">"
    html = html+"WebFont.load({"
    html = html+"google: {"
    html = html+"families: ["
    html = html+"\"Montserrat:100,100italic,200,200italic,300,300italic,400,400italic,500,500italic,600,600italic,700,700italic,800,800italic,900,900italic\","
    html = html+"    \"Lato:100,100italic,300,300italic,400,400italic,700,700italic,900,900italic\""
    html = html+"    ]"
    html = html+"  }"
    html = html+"});"
    html = html+"</script>"
    html = html+"<!--[if lt IE 9]><script src=\"https://cdnjs.cloudflare.com/ajax/libs/html5shiv/3.7.3/	html5shiv.min.js\" type=\"text/javascript\"></script><![endif]-->"
    html = html+"<script type=\"text/javascript\">"
    html = html+"! function (o, c) {"
    html = html+"  var n = c.documentElement,"
    html = html+"    t = \" w-mod-\";"
    html = html+"  n.className += t + \"js\", (\"ontouchstart\" in o || o.DocumentTouch && c instanceof DocumentTouch) && (n.className +="
    html = html+"\"t + \"touch\")"
    html = html+"}(window, document);"
    html = html+"</script>" 
    html = html+"</head>"
    html = html+"<body>"
    html = html+"<div class=\"container-2 w-container\" style=\"align-items: center !important;\">"
    html = html+"  <img src=\"https://empreendealagoas.com.br/src/img/logo-empreende-350.png\">"
    html = html+"</div>"
    html = html+"<div align=\"center\">"
    html = html+"<h1 class=\"heading\">Prezado aluno, seja bem vindo!</h1>"
    html = html+"<p class=\"paragraph\" style=\"text-align:center;\">Sua pré-matrícula no projeto Empreende Alagoas foi aceita.<br>"
    html = html+"Sua matrícula será liberada a partir do dia 25 de Novembro de 2021, data de início do projeto.<br>"
    html = html+"Esperamos que tenha um ótimo curso! Estaremos juntos nessa jornada! Conte conosco!"
    html = html+"  </p>"
    html = html+"</div>"  
    html = html+"</body>"
    html = html+"</html>"
    #message = "Prueba enviar mensaje"
    #-------------------------------------------
    #*Código para cada provedor de correo*
    #-------------------------------------------
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(html, 'html'))
    server.ehlo()
    #server.starttls()
    server.login(msg['From'], password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    print("successfull sent email to %s:" %(msg['To'])) 



def adiciona_anexo(msg, filename):
    if not os.path.isfile(filename):
        return

    ctype, encoding = mimetypes.guess_type(filename)

    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'

    maintype, subtype = ctype.split('/', 1)

    if maintype == 'text':
        with open(filename) as f:
            mime = MIMEText(f.read(), _subtype=subtype)
    elif maintype == 'image':
        with open(filename, 'rb') as f:
            mime = MIMEImage(f.read(), _subtype=subtype)
    elif maintype == 'audio':
        with open(filename, 'rb') as f:
            mime = MIMEAudio(f.read(), _subtype=subtype)
    else:
        with open(filename, 'rb') as f:
            mime = MIMEBase(maintype, subtype)
            mime.set_payload(f.read())

        encoders.encode_base64(mime)

    mime.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(mime)

def envia_email_anexo(email,message,subject):
    msg = MIMEMultipart()
    server = smtplib.SMTP_SSL('smtp.yandex.com.tr', 465)
    password = "123654bb"
    msg['From'] = "alex@wilivro.com.br"
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'html', 'utf-8'))
    adiciona_anexo(msg, arquivo)
    adiciona_anexo(msg, arquivo1)
    adiciona_anexo(msg, arquivo2)
    adiciona_anexo(msg, arquivo3)
    raw = msg.as_string()
    server.ehlo()
    #server.starttls()
    server.login(msg['From'], password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    print("successfull sent email to %s:" %(msg['To'])) 

def cria_turma(escola,projeto):
    sql = "select * from  escola where Nome = '"+escola+"'"
    cursor.execute(sql)
    escola = cursor.fetchone()
    if escola[0] is Null:
        print ("Escola :"+escola+" não existe")
        exit();
    id_escola = escola[0]
    sql = "select * from projeto where upper(Nome)=upper('"+projeto+"')"
    cursor.execute(sql)
    projeto = cursor.fetchone()
    if projeto[0] is Null:
        print ("Projeto :"+projeto+" não existe")
        exit();
    id_projeto = projeto[0]
    sql = "select * from projetoescola where IdProjeto='"+str(id_projeto)+"' and IdEscola ='"+str(id_escola)+"'"
    cursor.execute(sql)
    projeto_escola = cursor.fetchone()
    if projeto_escola[0] is Null:
        print ("Projeto escola :"+projeto+" não existe")
        exit();
    id_projeto_escola = projeto_escola[0]
    #disciplinas =["01-Empreendedorismo 2021","02-Mundo dos Micro e Pequenos Negócios 2021","03-Estratégias Digitais p/ Pequenos Empreendedores","04-Plano de Negócio 2021"]
    #pego projeto e as turmas
    for x, y in disciplinas.items():
        flag_turma = 0
        sql = "select p.id,p.nome,t.id,t.nome,pe.Id,d.Id from projeto p" 
        sql = sql+" join projetoescola pe on(pe.IdProjeto=p.id)" 
        sql = sql+" join escola e on(e.Id=pe.IdEscola)"
        sql = sql+" join turma t on(t.IdProjetoEscola=pe.id)"
        sql = sql+" join disciplina d on (d.Id=t.IdDisciplina)  where p.id="+str(id_projeto)
        sql = sql+ " and t.Nome like '"+x+"%' and d.Titulo='"+y+"'"
        sql = sql+" and e.id='"+str(id_escola)+"'  order by t.id"
        cursor.execute(sql)
        turmas = cursor.fetchall()
        if cursor.rowcount > 0:
            for turma in turmas:
                print("Verificando "+turma[3])
                sql = "select IdTurma,count(IdAluno) from alunoturma where IdTurma='"+str(turma[2])+"' having count(IdAluno)<20"
                cursor.execute(sql)
                matriculados = cursor.fetchone()
                if cursor.rowcount !=0:
                    flag_turma = 1
                nome_turma = turma[3]
                id_projeto_escola = turma[4]
                id_disciplina = turma[5]
            if flag_turma==0:
               end = None
               separador = "_" 
               pos_under = nome_turma.rfind(separador)
               if(pos_under == -1):
                   nome_turma = nome_turma+"_1"
               else:
                   pos_under = pos_under+1
                   numero_turma = nome_turma[pos_under:end] 
                   print(numero_turma)
                   nome_turma = x+"_"+str(int(numero_turma)+1)
               sql = "select * from turma where Nome='"+nome_turma+"' and IdProjetoEscola='"+str(id_projeto_escola)+"' and IdDisciplina='"+str(id_disciplina)+"'"
               print(sql)
               cursor.execute(sql)
               turma_existe = cursor.fetchone()
               if cursor.rowcount == 0:
                    sql = "insert into turma (MetodoAprovacao,Nome,IdProfessor,Status,IdProjetoEscola,BloqueiaMidia,IdDisciplina,Horario,"
                    sql = sql+"Domingo,Segunda,Terca,Quarta,Quinta,Sexta,Sabado,Capacidade,Estude,Pratique,Desempenho,Teste,PossuiSimulado,"
                    sql = sql+"PossuiRedacao,Minimo,TurmaAgrupada) values('1','"+nome_turma+"','145491','CADASTRADA','"+str(id_projeto_escola)+"','"
                    sql = sql+"0','"+str(id_disciplina)+"','08:00 ÁS 12:00','F','T','T','T','T','T','F','20','T','T','T','T','F','F','10','0')"
                    cursor.execute(sql)
                    print("Turma "+nome_turma+" foi criada")
               else:
                   print("turma :"+nome_turma+" ja existe")
                   exit()
        else:
            sql = "select * from disciplina where Titulo = '"+y+"'"
            cursor.execute(sql)
            disciplina = cursor.fetchone()
            if cursor.rowcount == 0:
                print("Não encontrei disciplina "+y)
                exit()
            nome_turma = x+"_1"
            id_disciplina = disciplina[0]
            sql = "insert into turma (MetodoAprovacao,Nome,IdProfessor,Status,IdProjetoEscola,BloqueiaMidia,IdDisciplina,Horario,"
            sql = sql+"Domingo,Segunda,Terca,Quarta,Quinta,Sexta,Sabado,Capacidade,Estude,Pratique,Desempenho,Teste,PossuiSimulado,"
            sql = sql+"PossuiRedacao,Minimo,TurmaAgrupada) values('1','"+nome_turma+"','144693','CADASTRADA','"+str(id_projeto_escola)+"','"
            sql = sql+"0','"+str(id_disciplina)+"','08:00 ÁS 12:00','F','T','T','T','T','T','F','20','T','T','T','T','F','F','10','0')"
            cursor.execute(sql)
    mydb.commit()

def matricula_aluno_turma(id_aluno,Rand_token,tipo):
    flag = 0
    if id_aluno ==0:
        print ("Id de aluno é inválido")
    else:
        nome_escola = "Escola Virtual - Empreendedor Alagoas"
        nome_projeto = "Empreende Alagoas"
        if tipo.find("Computador") != -1: #teste para curso on-line
            print ("Achei "+tipo)
            sql = "select * from  escola where Nome = '"+nome_escola+"'"
            cursor.execute(sql)
            escola = cursor.fetchone()
            id_escola =0;
            if cursor.rowcount == 0:
                sql = "insert into escola(Nome,IdEndereco,IdCliente)"
                sql = sql+" values('"+nome_escola+"',700,17)"
                print(sql)
                cursor.execute(sql)
                id_escola = cursor.lastrowid
                sql = "select * from projeto where upper(Nome)=upper('"+nome_projeto+"')" #projeto ja foi cadastrado no banco
                cursor.execute(sql)
                projeto = cursor.fetchone()
                sql = "insert into projetoescola (IdProjeto,IdEscola) values('"+str(projeto[0])+"','"+str(id_escola)+"')"
                cursor.execute(sql)
                id_projeto_escola = cursor.lastrowid
            else:
                id_escola = escola[0]
            print("Verificando turmas")
            cria_turma(nome_escola,nome_projeto)
        # disciplinas =["Empreendedorismo 2021","Mundo dos Micro e Pequenos Negócios 2021","Plano de Negócio 2021","Estratégias Digitais p/ Pequenos Empreendedores"]
            #pego projeto e as turmas
            sql = "select p.id,p.nome,t.id,t.nome,d.Id from projeto p" 
            sql = sql+" join projetoescola pe on(pe.IdProjeto=p.id)" 
            sql = sql+" join escola e on(e.Id=pe.IdEscola)"
            sql = sql+" join turma t on(t.IdProjetoEscola=pe.id)"
            sql = sql+" join disciplina d on (d.Id=t.IdDisciplina)  where upper(p.Nome)=upper('Empreende Alagoas')"
            sql = sql+" and e.id='"+str(id_escola)+"'"
            print(sql)
            cursor.execute(sql)
            turmas = cursor.fetchall()
            for turma in turmas:
                for x, y in disciplinas.items():
                    if turma[3].find(x) :
                        sql = "select count(*) from alunoturma where IdTurma='"+str(turma[2])+"'"
                        print(sql)
                        cursor.execute(sql)
                        matriculados = cursor.fetchone()
                        #print(matriculados)
                        if cursor.rowcount != 0 and matriculados[0] < 20:
                            print("Turma "+turma[3]+" tem vaga")
                            #sql = "select * from alunoturma where IdAluno='"+str(id_aluno)+"'"
                            #sql = sql+" and IdTurma='"+str(turma[2])+"'"
                            sql = "select atu.id,atu.IdAluno,t.nome,d.titulo,atu.data from alunoturma atu "
                            sql = sql+" join turma t on(t.id=atu.IdTurma)"
                            sql = sql+" join disciplina d on(d.id = t.IdDisciplina)"
                            sql = sql+" where atu.IdAluno='"+str(id_aluno)+"' and d.id='"+str(turma[4])+"'"
                            print(sql)
                            cursor.execute(sql)
                            existe = cursor.fetchall()
                            numero_cursos = cursor.rowcount
                            if cursor.rowcount == 0 or existe is Null:
                                sql = "insert into alunoturma (IdAluno,IdTurma,Data)"
                                sql = sql+" values('"+str(id_aluno)+"','"+str(turma[2])+"','"+data_e_hora_em_texto+"')"
                                print(sql)
                                cursor.execute(sql)
                                flag=1
                            else:
                                if numero_cursos > 1:
                                    idAlunoturma =0
                                    for num in (existe):
                                        if num[0] > idAlunoturma:
                                            idAlunoturma = num[0]
                                    sql = "select * from respostaaluno where IdAlunoTurma="+str(idAlunoturma)
                                    print(sql)
                                    cursor.execute(sql)
                                    cursor.fetchone()
                                    respostaaluno = cursor.rowcount
                                    sql = "select * from alunoturmaedc where IdAlunoTurma="+str(idAlunoturma)
                                    print(sql)
                                    cursor.execute(sql)
                                    cursor.fetchone()
                                    alunoturmaedc = cursor.rowcount
                                    if respostaaluno == 0 and alunoturmaedc == 0:
                                        sql = "delete from alunoturma where id="+str(idAlunoturma)
                                        print(sql)
                                        cursor.execute(sql)
                                        print("removi matricula "+str(idAlunoturma))
                                    else:
                                        print("Matricula "+str(idAlunoturma)+" ja foi utilizada")
                                    
                        else:
                            print("Turma "+turma[3]+" esta cheia")
                            
            if flag==1:
                sql = "select Nome,Email,reqCode from usuario where id='"+str(id_aluno)+"'"
                cursor.execute(sql)
                aluno = cursor.fetchone()
                if cursor.rowcount !=0:
                    print(aluno)
                    link = "https://wiquadro.com.br/site/alterarsenha/validationCode/"+aluno[2]
                    matriculas_feitas.writerow([aluno[0], aluno[1],link])
                    #planMatriculados.append_row([aluno[0], aluno[1],link])
                  
        else:
            print("Curso presencial")


                    


# fim área de funções
#----------------------------------------------------------------------------------------------------------------------------

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)


# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("#EmpreendeAlagoas - pré-matrícula (respostas)").sheet1
#planMatriculados = client.open("matriculados_empreendealagoas").sheet1

# Extract and print all of the values
list_of_hashes = sheet.get_all_values()
# i=1
#print(sheet.row_values(1))
#exit()

#arquivo csv de saida
now = datetime.now()
data_atual = now.strftime("%d_%m_%Y_%H_%M_%S")
arquivo = data_atual+"_inscritos_empreendedor_alagoas.csv"
arquivo1 = data_atual+"_erros_empreendedor_alagoas.csv"
arquivo2 = data_atual+"_inscricoes_ja_existentes_empreendedor_alagoas.csv"
arquivo3 = data_atual+"_matriculados_empreendedor_alagoas.csv"
f = open(arquivo, 'w', newline='', encoding='utf-8')
e = open(arquivo1, 'w', newline='', encoding='utf-8')
r = open(arquivo2, 'w', newline='', encoding='utf-8')
m = open(arquivo3, 'w', newline='', encoding='utf-8')
w = csv.writer(f)
erro = csv.writer(e)
existente = csv.writer(r)
matriculas_feitas = csv.writer(m)
matriculas_feitas.writerow(['Nome completo:', 'Email:','Troca Senha'])
w.writerow(['Nome completo:', 'Situação acadêmica:', 'Telefone:', 'Email:', 'Cidade onde mora:', 'Como irá fazer o curso: ', 'Telecentros disponíveis:', 'CPF', 'CNPJ', 'Data de nascimento', 'Endereço:'])
erro.writerow(['Erro','Nome completo:', 'Situação acadêmica:', 'Telefone:', 'Email:', 'Cidade onde mora:', 'Como irá fazer o curso: ', 'Telecentros disponíveis:', 'CPF', 'CNPJ', 'Data de nascimento', 'Endereço:'])
existente.writerow(['Nome completo:', 'Situação acadêmica:', 'Telefone:', 'Email:', 'Cidade onde mora:', 'Como irá fazer o curso: ', 'Telecentros disponíveis:', 'CPF', 'CNPJ', 'Data de nascimento', 'Endereço:'])
for i in range(len(list_of_hashes)):
    if i <= 1000:
        print(str(i))
        if validate_cpf(list_of_hashes[i][8]):
            print (list_of_hashes[i][8])
            if list_of_hashes[i][4] != '':
                sql = "SELECT * FROM usuario WHERE Email='"+list_of_hashes[i][4]+"'"
                sql = sql+" or cpf='"+list_of_hashes[i][8]+"'"
                cursor.execute(sql)
                myresult = cursor.fetchone()
                if cursor.rowcount != 0 :
                    #print(myresult)
                    if myresult[8] is not None and myresult[8] !='':
                        Rand_token = myresult[8]
                    else:
                        Rand_token = uuid4()
                        sql = "UPDATE usuario set reqCode='"+str(Rand_token)+"', Email='"+list_of_hashes[i][4]+"',"
                        sql=sql+" cpf='"+list_of_hashes[i][8]+"' WHERE Id='"+str(myresult[0])+"'"
                        cursor.execute(sql)
                        sql = "select * from usuariogrupousuario WHERE IdUsuario='"+str(myresult[0])+"'"
                        cursor.execute(sql)
                        usuario_grupo = cursor.fetchall()
                        if cursor.rowcount == 0 :
                            sql = "INSERT INTO usuariogrupousuario(IdUsuario,IdGrupoUsuario) VALUES"
                            sql = sql+" ('"+str(myresult[0])+"','4')"
                            cursor.execute(sql)

                    existente.writerow([list_of_hashes[i][1],list_of_hashes[i][2],list_of_hashes[i][3],list_of_hashes[i][4],list_of_hashes[i][5],list_of_hashes[i][6],list_of_hashes[i][7],list_of_hashes[i][8],list_of_hashes[i][9],list_of_hashes[i][10],list_of_hashes[i][11]])
                    #print(sql)
                    #print(myresult)
                    matricula_aluno_turma(myresult[0],Rand_token,list_of_hashes[i][6])
                else:
                    Rand_token = uuid4()
                    password = list_of_hashes[i][8]
                    password = hashlib.md5(password.encode())
                    data_aux = datetime.strptime(list_of_hashes[i][10],'%d/%m/%Y')
                    data_nascimento = data_aux.strftime('%Y-%m-%d')
                    sql = "INSERT INTO usuario(Nome,Email,Nascimento,Cidade,CPF,reqCode,senha,uf,UniqId) VALUES"
                    sql = sql+" ('"+list_of_hashes[i][1]+"','"+list_of_hashes[i][4]+"','"+data_nascimento
                    sql = sql+"','"+list_of_hashes[i][5]+"','"+list_of_hashes[i][8]+"','"+str(Rand_token)+"','"+password.hexdigest()+"','AL','{"+str(Rand_token)+"}')"
                    cursor.execute(sql)
                    if cursor.lastrowid is not Null:
                        id_usuario = cursor.lastrowid
                        sql = "INSERT INTO usuariogrupousuario(IdUsuario,IdGrupoUsuario) VALUES"
                        sql = sql+" ('"+str(id_usuario)+"','4')"
                        cursor.execute(sql)
                        #print(cursor.lastrowid)
                        matricula_aluno_turma(id_usuario,Rand_token,list_of_hashes[i][6])
                        w.writerow([list_of_hashes[i][1],list_of_hashes[i][2],list_of_hashes[i][3],list_of_hashes[i][4],list_of_hashes[i][5],list_of_hashes[i][6],list_of_hashes[i][7],list_of_hashes[i][8],list_of_hashes[i][9],list_of_hashes[i][10],list_of_hashes[i][11]])
                    else:
                        print ("Erro em "+sql) 
                        #exit()
            mydb.commit()      
            #envia_email_aluno_cadastrado('alexmundi@gmail.com',mensagem,'Bem vindo ao empreendedor alagoas')
            #exit()'''
        else:
            erro.writerow(['CPF Inválido',list_of_hashes[i][1],list_of_hashes[i][2],list_of_hashes[i][3],list_of_hashes[i][4],list_of_hashes[i][5],list_of_hashes[i][6],list_of_hashes[i][7],list_of_hashes[i][8],list_of_hashes[i][9],list_of_hashes[i][10],list_of_hashes[i][11]])
        if i%100 == 0:
            cursor.reset()
            cursor.close()
            cursor = mydb.cursor(buffered=True)
    else:
        print("inclusão cancelada")

f.close() 
e.close()
r.close()
m.close()
mydb.commit()
print("terminei "+ str(i))
#envia_email_anexo("jordana@wilivro.com.br","Arquivos .csv de hoje","Inscritos Empreendedor Alagoas")
#print(sheet.row_values(1))
#res = sheet.row_values(100)

#print(res[8])
