import time
import mysql.connector
from  mysql.connector import Error
import win32com.client as win32
import pandas as pd     

def email(grupo,emails):

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = emails
    email.Subject = "Email Automático bases atualizadas"
    email.HTMLBody = f"""

        <!--Start Three Blocks-->
        <table width='440' border='0' cellpadding='0' cellspacing='0' align='center' class='deviceWidth'>
          <tr>
            <td width='100%' bgcolor='#fff' style='border-radius: 4px; 
                                                   -webkit-box-shadow: 0px 0px 15px 15px rgba(0,0,0,0.15);
                                                   -moz-box-shadow: 0px 0px 15px 15px rgba(0,0,0,0.15);
                                                   box-shadow: 0px 0px 15px 15px rgba(0,0,0,0.15);' >
            <!-- Top  -->
              <table width='100%'  border='0' cellpadding='0' cellspacing='0' align='center' > 
              <tr>
                  <td  width='50%' class='center' style='font-size: 20px; 
                                                        color: #000000; 
                                                        font-weight: normal; 
                                                        text-align: center; 
                                                        font-family: Arial, Helvetica, sans-serif; 
                                                        line-height: 10px; 
                                                        vertical-align: middle; 
                                                        padding:20px 10px;  '>
                    <p>  Consolidado relatórios atualizados!</p>
                    <p>  {grupo}  </p>
                  </td>
                </tr>           
        
                <tr class='bg-color'>
                  <td width='50%' class='center' style='font-size: 20px; color: #000000; font-weight: normal; text-align: center; font-family: Arial, Helvetica, sans-serif; line-height: 25px; vertical-align: middle; padding: 5px 10px 20px 10px;' >
                    <span><p>Equipe á disposição! </p></span>
          
                 </td>
               </tr>
        
               
            </table>
    """

    # Anexar arquivo
    anexo = "C:/Users/Documents/script_automático_tratativa_e_email/uploads/email/Consolidado.xlsx"
    email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviador")   

def deu_errado_email(erro,buscar,grupo):

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = "email_ADM_@teste.br"
    email.Subject = "Informação sobre atualização da base..."
    email.HTMLBody = f"""

        <!--Start Three Blocks-->
        <table width='440' border='0' cellpadding='0' cellspacing='0' align='center' class='deviceWidth'>
          <tr>
            <td width='100%' bgcolor='#fff' style='border-radius: 4px; 
                                                   -webkit-box-shadow: 0px 0px 15px 15px rgba(0,0,0,0.15);
                                                   -moz-box-shadow: 0px 0px 15px 15px rgba(0,0,0,0.15);
                                                   box-shadow: 0px 0px 15px 15px rgba(0,0,0,0.15);' >
            <!-- Top  -->
              <table width='100%'  border='0' cellpadding='0' cellspacing='0' align='center' > 
              <tr>
                  <td  width='50%' class='center' style='font-size: 20px; 
                                                        color: #687074; 
                                                        font-weight: normal; 
                                                        text-align: center; 
                                                        font-family: Arial, Helvetica, sans-serif; 
                                                        line-height: 10px; 
                                                        vertical-align: middle; 
                                                        padding:20px 10px;  '>
                    <p> VERIFICAR ROBÔ DO CONSOLIDADO DE RELATÓRIOS ATUALIZADOS POR GRUPO </p>
                  </td>
                </tr>
        
             <tr class='bg-color'>
                  <td width='50%' class='center' style='font-size: 20px; color: #687074; font-weight: normal; text-align: center; font-family: Arial, Helvetica, sans-serif; line-height: 25px; vertical-align: middle; padding: 5px 10px 20px 10px;' >
                    <span><p> ERROR = {erro} na parte de {buscar}. </p></span>
          
                 </td>
               </tr>
                <tr class='bg-color'>
                  <td width='50%' class='center' style='font-size: 20px; color: #687074; font-weight: normal; text-align: center; font-family: Arial, Helvetica, sans-serif; line-height: 25px; vertical-align: middle; padding: 5px 10px 20px 10px;' >
                    <span><p> GRUPO =  {grupo}. </p></span>
          
                 </td>
               </tr>
            
        
                <tr class='bg-color'>
                  <td width='50%' class='center' style='font-size: 20px; color: #687074; font-weight: normal; text-align: center; font-family: Arial, Helvetica, sans-serif; line-height: 25px; vertical-align: middle; padding: 5px 10px 20px 10px;' >
                    <span><p>Equipe á disposição! </p></span>
          
                 </td>
               </tr>
        
               
            </table>
    """

    email.Send()
    print("Email Enviador")
    
def bucar_dados_do_banco(grupo):
    

    try:
       
        con = mysql.connector.connect (host='', database='',
        user='', password= '')
        consulta =  "{}{}{}{}".format("select * from TABELA where DATE_FORMAT(data, '%Y-%m-%d') = CURDATE() and grupo = ","'",grupo,"'")
        consulta_sql = consulta
        cursor = con.cursor()
        cursor.execute(consulta_sql)
        linhas= cursor.fetchall()
        print ("Número total de registros retornados:", cursor.rowcount)
        for linha in linhas:
            nome.append(linha[4])
            grupo_.append(linha[5])
            data.append("{:%d/%m/%Y}".format(linha[3]))
        #===================================================================           
        gerar_excel(nome,grupo_,data)
        time.sleep(30)
        #===================================================================
        email(grupo,emails)
           #===================================================================
        nome.clear()
        grupo_.clear()
        data.clear()
    except Error as erro:
        print("Erro ao acessar tabela MysQL", erro)
        #===================================================================
        deu_errado_email(erro,'buscar dados',grupo)
    finally:
        if(con.is_connected()):
            con.close()
            cursor.close()
            print("Conexäo ao MysQL encerrada")

def gerar_excel(nome,grupo_,data):

    lista_de_tuplas = list(zip(nome, grupo_, data))
    # converte uma lista de tuplas num DataFrame
    df = pd.DataFrame(lista_de_tuplas, columns=['Nome', 'grupo_','Data'])
    df.to_excel(r'uploads/email/Consolidado.xlsx')



nome = []
grupo_= []
data = []
# ===========================================================================================

grupo ='grupo1'
emails = 'email_de_teste@teste.br;teste@teste.br;teste3@teste.br'
bucar_dados_do_banco(grupo)
time.sleep(60)

grupo = 'grupo2'
emails = 'email_de_teste@teste.br;teste@teste.br;teste3@teste.br'
bucar_dados_do_banco(grupo)
time.sleep(60)

grupo ='grupo3'
emails = 'email_de_teste@teste.br;teste@teste.br;teste3@teste.br'
bucar_dados_do_banco(grupo)
time.sleep(60)

print("=====================FIM=====================")