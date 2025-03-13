import os
import pandas as pd
import re
from fuzzywuzzy import process
from datetime import datetime
import textwrap

# Caminho para a pasta "resultados"
diretorio_resultados = os.path.join(os.getcwd(), 'resultados')
diretorio_processos_sief = os.path.join(os.getcwd(), 'processos sief')
# Caminhos para as pastas e arquivos
diretorio_codigos = os.path.join(os.getcwd(), 'resultados_codigos')
arquivo_tabelas = os.path.join(os.getcwd(), 'TABELASCDIGOSDERECEITA.xlsx')
tabela_depto_pessoal = pd.read_excel(arquivo_tabelas, sheet_name='Depto Pessoal')
tabela_fiscal = pd.read_excel(arquivo_tabelas, sheet_name='Fiscal')
diretorio_gps = os.path.join(os.getcwd(), 'resultado_gfip_gps')


def salvar_mensagem(df_existente, nome_empresa, nova_mensagem, caminho_saida):
    # Lista de empresas já existentes no arquivo
    nomes_existentes = df_existente['Empresa'].tolist()

    # Mostrar os nomes das empresas existentes no DataFrame
    print("Empresas existentes no arquivo:", nomes_existentes)

    # Encontrar o nome mais parecido
    nome_mais_proximo, score = process.extractOne(nome_empresa, nomes_existentes) if nomes_existentes else (None, 0)

    # Mostrar o nome mais próximo e o score
    print(f"Procurando pelo nome: {nome_empresa}")
    print(f"Nome mais próximo encontrado: {nome_mais_proximo}, Score: {score}")

    # Se encontrou uma correspondência confiável, usa o nome existente
    if nome_mais_proximo and score >= 80:
        nome_empresa = nome_mais_proximo
        print(f"Usando nome mais próximo: {nome_empresa}")

    # Se a empresa já existir no arquivo, concatena a mensagem
    if nome_empresa in df_existente['Empresa'].values:
        print(f"Empresa '{nome_empresa}' encontrada no arquivo, concatenando mensagem...")
        df_existente.loc[df_existente['Empresa'] == nome_empresa, 'Mensagem'] += f"\n{nova_mensagem}"
    else:
        print(f"Empresa '{nome_empresa}' não encontrada no arquivo, criando nova linha...")
        nova_linha = pd.DataFrame({"Empresa": [nome_empresa], "Mensagem": [nova_mensagem]})
        df_existente = pd.concat([df_existente, nova_linha], ignore_index=True)

    return df_existente


def criar_msgs_gps(diretorio_gps, caminho_saida):
    data_atual = datetime.now().strftime("%d/%m/%y")

    if os.path.exists(caminho_saida):
        df_existente = pd.read_excel(caminho_saida)
    else:
        df_existente = pd.DataFrame(columns=["Empresa", "Mensagem"])

    for arquivo in os.listdir(diretorio_gps):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            caminho_arquivo = os.path.join(diretorio_gps, arquivo)

            # Extraindo CNPJ ou CPF do nome do arquivo
            match = re.match(r'(\d{11}|\d{14})_(.+)\.xls[x]?', arquivo)
            if not match:
                print(f"⚠️ Nome de arquivo inválido: {arquivo}")
                continue

            cnpj_cpf, nome_empresa = match.groups()

            print(f"🔍 Processando {nome_empresa} - {cnpj_cpf}")

            df = pd.read_excel(caminho_arquivo)

            if {'Valores', 'PA - EXERC.'}.issubset(df.columns):
                # Remover linhas vazias
                df = df.dropna(subset=['Valores', 'PA - EXERC.'])

                # Ajustar formato do PA - Exercício (para garantir consistência)
                def formatar_pa_exercicio(pa_exercicio):
                    try:
                        pa_exercicio_str = str(pa_exercicio).strip()
                        if len(pa_exercicio_str.split('/')) == 3:
                            return '/'.join(pa_exercicio_str.split('/')[1:])
                        if re.match(r"^\d{2}/\d{4}$", pa_exercicio_str):
                            return pa_exercicio_str
                        return pa_exercicio_str
                    except Exception:
                        return None

                df['PA - EXERC.'] = df['PA - EXERC.'].apply(formatar_pa_exercicio)

                # Agrupar valores por PA - Exercício
                debitos_por_pa = df.groupby('PA - EXERC.')['Valores'].sum().to_dict()

                # Construir a mensagem
                mensagem = f"\n\n{nome_empresa}, identificamos que sua empresa possui débitos de INSS na Receita Federal.\n"
                mensagem += "Segue o resumo dos valores pendentes:\n\n"

                for pa_exercicio, valor in debitos_por_pa.items():
                    print('Valorrrrrrrrrr: ', valor,'\n')
                    if isinstance(valor, str):
                        valor = valor.replace('.', '').replace(',', '.').strip()
                    
                    try:
                        valor = float(valor)
                    except ValueError:
                        valor = 0.0
                    mensagem += f"  - Referente a {pa_exercicio}: R$ {valor:.2f}\n"

                mensagem += "\nEsses débitos fazem parte do **Departamento Pessoal**.\n"

                df_existente = salvar_mensagem(df_existente, cnpj_cpf, mensagem.strip(), caminho_saida)

                print(f"Mensagem gerada para {nome_empresa}:\n{mensagem}\n")
            else:
                print(f"⚠️ O arquivo {arquivo} não possui as colunas esperadas.")

    df_existente.to_excel(caminho_saida, index=False)
    print("Mensagens salvas com sucesso!")


def criar_msgs(caminho_saida):
    data_atual = datetime.now().strftime("%d/%m/%y")
    
    # Percorre todos os arquivos Excel na pasta
    for arquivo in os.listdir(diretorio_resultados):

        if os.path.exists(caminho_saida):
            df_existente = pd.read_excel(caminho_saida)
        else:
            df_existente = pd.DataFrame(columns=["Empresa", "Mensagem"])

        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):  # Verifica se é um arquivo Excel
            caminho_arquivo = os.path.join(diretorio_resultados, arquivo)
            
            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            
            # Garante que as colunas necessárias estão no DataFrame
            if {'EMPRESA', 'DÍVIDA ATIVA', 'NUMERO DO PROCESSO', 'SITUAÇÃO'}.issubset(df.columns):
                
                # Tenta extrair o CNPJ limpo (14 dígitos) da coluna "EMPRESA"
                cnpj = re.search(r'(\d{14})', str(df['EMPRESA'].iloc[0]))  # Supondo que o CNPJ esteja na primeira linha
                if cnpj:
                    cnpj = cnpj.group(1)  # Extrai o CNPJ limpo
                    
                    # Remover o CNPJ do nome da empresa para utilizá-lo na mensagem
                    nome_empresa_sem_cnpj = df['EMPRESA'].iloc[0].replace(cnpj + "_", "")  # Remove o CNPJ do início do nome
                    
                    print(f"🔍 Buscando pelo CNPJ: {cnpj}")
                    
                    # Agrupa os processos pela mesma situação
                    situacoes = df.groupby('SITUAÇÃO')['NUMERO DO PROCESSO'].apply(list).to_dict()
                    
                    # Gera a mensagem personalizada para a empresa (usando o nome sem o CNPJ)
                    mensagem = f"A empresa {nome_empresa_sem_cnpj} possui os seguintes débitos na Procuradoria-Geral da Fazenda Nacional: \n"
                    for situacao, processos in situacoes.items():
                        processos_formatados = ', '.join(map(str, processos))  # Junta os números dos processos
                        mensagem += f"{situacao}'.\n"

                    df_existente = salvar_mensagem(df_existente, cnpj, mensagem.strip(), caminho_saida)
                    
                    print(f"Mensagem para {nome_empresa_sem_cnpj}:\n{mensagem}\n")
                else:
                    print(f"⚠️ CNPJ não encontrado para a empresa '{df['EMPRESA'].iloc[0]}'.")
            else:
                print(f"O arquivo {arquivo} não possui as colunas esperadas.")
            
        df_existente.to_excel(caminho_saida, index=False)
        print("Mensagens salvas com sucesso!")



def criar_msgs_processos_sief(caminho_saida, diretorio_processos_sief):
    from datetime import datetime
    import os
    import pandas as pd

    data_atual = datetime.now().strftime("%d/%m/%y")
    
    # Verifica se já existe um arquivo com mensagens e carrega os dados
    if os.path.exists(caminho_saida):
        df_existente = pd.read_excel(caminho_saida)
    else:
        df_existente = pd.DataFrame(columns=["Empresa", "Mensagem"])

    # Percorre todos os arquivos Excel na pasta
    for arquivo in os.listdir(diretorio_processos_sief):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            # Exemplo do nome do arquivo: "23098061000139_C R V ESTERO E CIA LTDA.xlsx"
            # Extrai o CNPJ que está antes do primeiro '_'
            cnpj = arquivo.split('_')[0]
            print(f"🔍 CNPJ extraído do nome do arquivo: {cnpj}")

            # Monta o caminho completo do arquivo
            caminho_arquivo = os.path.join(diretorio_processos_sief, arquivo)
            
            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            
            # Verifica se a coluna necessária existe
            if 'Processos SIEF' in df.columns:
                # Lista de processos (removendo valores vazios)
                processos = df["Processos SIEF"].dropna().astype(str).tolist()

                match = re.match(r'^\d+_(.*)\.xlsx$', arquivo)
                if match:
                    nome_empresa_sem_cnpj = match.group(1)
                else:
                    nome_empresa_sem_cnpj = "Nome não encontrado"
                
                # Gera a mensagem personalizada usando somente o CNPJ como identificador
                mensagem = f"\n\nA empresa {nome_empresa_sem_cnpj} possui os seguintes débitos referentes a Processos SIEF:\n\n"
                mensagem += ', '.join(processos)
                
                # Salva ou concatena a mensagem no DataFrame existente, usando o CNPJ como chave
                df_existente = salvar_mensagem(df_existente, cnpj, mensagem.strip(), caminho_saida)
                
                print(f"✅ Mensagem gerada para {cnpj}:\n{mensagem}\n")
            else:
                print(f"⚠️ O arquivo {arquivo} não possui a coluna 'Processos SIEF' esperada.")

    # Salva as mensagens geradas no arquivo Excel
    df_existente.to_excel(caminho_saida, index=False)
    print("✅ Mensagens salvas com sucesso!")

#criar_msgs_processos_sief(caminho_saida='mensagens.xlsx',diretorio_processos_sief=diretorio_processos_sief)



def criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal, caminho_saida):
    
    data_atual = datetime.now().strftime("%d/%m/%y")
    if os.path.exists(caminho_saida):
        df_existente = pd.read_excel(caminho_saida)
    else:
        df_existente = pd.DataFrame(columns=["Empresa", "Mensagem"])

    for arquivo in os.listdir(diretorio_codigos):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            caminho_arquivo = os.path.join(diretorio_codigos, arquivo)
            df = pd.read_excel(caminho_arquivo)

            if {'Empresa', 'Código Fiscal', 'PA - Exercício', 'Saldo Devedor Consignado'}.issubset(df.columns):
                # Tenta extrair o CNPJ limpo (14 dígitos) da coluna "Empresa"
                cnpj = re.search(r'(\d{14})', str(df['Empresa'].iloc[0]))  # Supondo que o CNPJ esteja na primeira linha
                if cnpj:
                    cnpj = cnpj.group(1)  # Extrai o CNPJ limpo
                    
                    # Remover o CNPJ do nome da empresa para utilizá-lo na mensagem
                    nome_empresa_sem_cnpj = df['Empresa'].iloc[0].replace(cnpj + "_", "")  # Remove o CNPJ do início do nome
                    
                    print(f"🔍 Buscando pelo CNPJ: {cnpj}")
                    
                    # Função para ajustar o formato do PA - Exercício
                    def formatar_pa_exercicio(pa_exercicio):
                        try:
                            pa_exercicio_str = str(pa_exercicio).strip()
                            # Se estiver no formato DDD/MM/YYYY, remove o dia e retorna MM/AAAA
                            if len(pa_exercicio_str.split('/')) == 3:
                                return '/'.join(pa_exercicio_str.split('/')[1:])
                            # Se for apenas um número seguido de "º", adiciona " TRIMESTRE"
                            if re.match(r"^\dº$", pa_exercicio_str):
                                return f"{pa_exercicio_str} TRIMESTRE"
                            # Se estiver no formato MM/AAAA, retorna como está
                            if re.match(r"^\d{2}/\d{4}$", pa_exercicio_str):
                                return pa_exercicio_str
                            return pa_exercicio_str
                        except Exception as e:
                            print(f"Erro ao formatar PA - Exercício: {pa_exercicio}, erro: {e}")
                            return None

                    df['PA - Exercício'] = df['PA - Exercício'].apply(formatar_pa_exercicio)

                    # Agrupando os dados pelo PA - Exercício
                    meses_agrupados = df.groupby('PA - Exercício')

                    mensagem = f"Olá {nome_empresa_sem_cnpj},\n"
                    mensagem += "Identificamos que sua empresa possui algumas pendências em aberto junto à Receita Federal.\n"
                    mensagem += "Essas pendências podem gerar multas, juros e complicações mais sérias se não forem regularizadas em tempo hábil.\n\n"
                    mensagem += "Segue o resumo dos seus débitos:\n\n"
                    mensagem = textwrap.dedent(mensagem)


                    for pa_exercicio, grupo in meses_agrupados:
                        mensagem += f"**Referente a {pa_exercicio}:**\n"
                        debitos_por_tipo = {}

                        for _, row in grupo.iterrows():
                            codigo_fiscal_completo = str(row['Código Fiscal']).strip()
                            saldo_devedor = str(row['Saldo Devedor Consignado']).replace('.', '').replace(',', '.')
                            saldo_devedor = float(saldo_devedor)

                            try:
                                saldo_devedor = float(saldo_devedor)
                            except ValueError:
                                saldo_devedor = 0.0  # Caso o valor não seja numérico, considera como zero

                            if saldo_devedor <= 0:
                                continue  # Ignora débitos zerados
                            
            
                            match = re.match(r'(\d+)[-/](\d+)', codigo_fiscal_completo)
                            if match:
                                codigo_fiscal_formatado_original = f"{match.group(1)}-{match.group(2)}"
                                codigo_fiscal_com_variacao = f"{match.group(1)}/{match.group(2)}"
                            else:
                                codigo_fiscal_formatado_original = codigo_fiscal_completo
                                codigo_fiscal_com_variacao = codigo_fiscal_completo

                            # Verifica em qual tabela o código está presente
                            
                            descricao = re.sub(r'^\d+[-/]\d+\s-\s', '', codigo_fiscal_completo).strip()

                            # Se a descrição tiver PIS ou COFINS, define como Fiscal
                            if "PIS" in descricao.upper() or "COFINS" in descricao.upper():
                                tipo_debito = "Departamento Fiscal"
                            elif (codigo_fiscal_formatado_original in tabela_depto_pessoal['Código de receita'].astype(str).values or
                                codigo_fiscal_com_variacao in tabela_depto_pessoal['Código de receita'].astype(str).values):
                                tipo_debito = "Departamento Pessoal"
                            elif (codigo_fiscal_formatado_original in tabela_fiscal['Código de receita'].astype(str).values or
                                codigo_fiscal_com_variacao in tabela_fiscal['Código de receita'].astype(str).values):
                                tipo_debito = "Fiscal"
                            else:
                                tipo_debito = f"outros ({descricao})"


                            print(f"PA: {pa_exercicio}, Código: {codigo_fiscal_completo}, Tipo: {tipo_debito}, Valor: {saldo_devedor}")

                            # Soma os valores por tipo de débito
                            if tipo_debito in debitos_por_tipo:
                                debitos_por_tipo[tipo_debito] += saldo_devedor
                            else:
                                debitos_por_tipo[tipo_debito] = saldo_devedor

                        # Adiciona os valores somados à mensagem
                        for tipo, valor in debitos_por_tipo.items():
                            mensagem += f"  - {tipo}: R$ {valor:.2f}\n"
                        
                        mensagem += "\n"  # Separação entre meses

                    df_existente = salvar_mensagem(df_existente, cnpj, mensagem.strip(), caminho_saida)
                    print(f"Mensagem gerada para {nome_empresa_sem_cnpj}:\n{mensagem}\n")
                else:
                    print(f"⚠️ CNPJ não encontrado para a empresa '{df['Empresa'].iloc[0]}'.")
            else:
                print(f"O arquivo {arquivo} não possui as colunas esperadas.")

    df_existente.to_excel(caminho_saida, index=False)
    print("Mensagens salvas com sucesso!")


# Chamada da função
import pandas as pd

def criar_msg_fgts():
    # Carregar os arquivos
    fgts_df = pd.read_excel("debitos_fgts.xlsx")
    mensagens_df = pd.read_excel("mensagens.xlsx")

    # Criar um dicionário para agrupar os débitos por empresa
    fgts_dict = {}
    for _, row in fgts_df.iterrows():
        nome_completo = row["Nome da Empresa"]
        cnpj, nome_empresa = nome_completo.split("_", 1)
        mes_ref = row["Mês Ref."]
        valor = row["Valor Débitos"]
        
        if cnpj not in fgts_dict:
            fgts_dict[cnpj] = {"nome": nome_empresa, "debitos": {}}
        
        if mes_ref not in fgts_dict[cnpj]["debitos"]:
            fgts_dict[cnpj]["debitos"][mes_ref] = 0
        
        fgts_dict[cnpj]["debitos"][mes_ref] += valor

    # Criar ou atualizar as mensagens
    for cnpj, data in fgts_dict.items():
        nome_empresa = data["nome"]
        debitos_texto = ", ".join([f"{mes}: R$ {valor:.2f}" for mes, valor in data["debitos"].items()])
        
        if cnpj in mensagens_df["Empresa"].astype(str).values:
            print('tinha o cnpj', cnpj)
            mensagem_fgts = f"{nome_empresa}, você também possui débitos de FGTS: " + ", ".join(
                [f"{mes} no valor de R$ {valor:.2f}" for mes, valor in data['debitos'].items()]
            ) + "."
            mensagens_df.loc[mensagens_df["Empresa"].astype(str) == cnpj, "Mensagem"] += f" {mensagem_fgts}"
        else:
            mensagem = f"{nome_empresa}, segue resumo dos seus débitos de FGTS: {debitos_texto}."
            mensagens_df = pd.concat([mensagens_df, pd.DataFrame({"Empresa": [cnpj], "Mensagem": [mensagem]})], ignore_index=True)

    # Salvar o arquivo atualizado
    mensagens_df.to_excel("mensagens.xlsx", index=False)

    print("Mensagens de FGTS geradas e salvas com sucesso!")


def criar_msg_final():
    # Carregar o arquivo de mensagens
    mensagens_df = pd.read_excel("mensagens.xlsx")

    # Definir a mensagem final
    data_atual = datetime.now().strftime("%d/%m/%y")
    mensagem_final = (
        f"\nOs valores informados são válidos na data de envio deste e-mail ({data_atual}) e podem sofrer alterações.\n"
        "Caso tenha interesse em regularizar essas pendências, entre em contato com o nosso time "
        "para mais detalhes e orientações sobre os próximos passos.\n"
        "Ficamos à disposição para qualquer dúvida ou informação adicional!\n\n"
        "Atenciosamente,\n"
        "Prímor Contábil\n"
        "(44) 98462-9927 / atendimento@contabilprimor.com.br"
    )

    # Garantir que a mensagem final seja a última coisa adicionada a cada linha
    mensagens_df["Mensagem"] = mensagens_df["Mensagem"].astype(str) + mensagem_final

    # Salvar as mensagens atualizadas
    mensagens_df.to_excel("mensagens.xlsx", index=False)

    print("Mensagem final adicionada com sucesso!")


#criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal, caminho_saida = 'mensagens.xlsx')

def criar_msgs_geral():
    criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal, caminho_saida = 'mensagens.xlsx')
    criar_msgs(caminho_saida="mensagens.xlsx")
    criar_msgs_processos_sief(caminho_saida="mensagens.xlsx", diretorio_processos_sief = diretorio_processos_sief)
    criar_msgs_gps(diretorio_gps,caminho_saida='mensagens.xlsx')
    #criar_msg_fgts()
    criar_msg_final()

criar_msgs_geral()