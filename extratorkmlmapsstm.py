import os
import streamlit as st
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import openpyxl
import csv
from collections import defaultdict
import webbrowser

class KMLProcessorApp:
    def __init__(self):
        self.arquivo_kml = None
        self.arquivo_db = None
        self.arquivo_db_cto = None
        self.dados_db = defaultdict(list)

    def carregar_kml(self):
        arquivo_kml = st.file_uploader(
            "Escolha o arquivo KML ou KMZ",
            type=["kml", "kmz"]
        )
        if arquivo_kml:
            self.arquivo_kml = arquivo_kml
            st.write(f"Arquivo KML carregado: {arquivo_kml.name}")

    def carregar_db(self):
        arquivo_db = st.file_uploader(
            "Escolha o arquivo do banco de dados Clientes (Excel)",
            type=["xlsx"]
        )
        if arquivo_db:
            self.arquivo_db = arquivo_db
            wb = openpyxl.load_workbook(arquivo_db)
            sheet = wb.active
            for excel_row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True):
                nome_tratado = excel_row[0].split(' ')[0]
                ids = excel_row[6]
                status = excel_row[4]
                self.dados_db[nome_tratado].append({
                    'id': ids,
                    'status': status
                })
            st.write(f"Arquivo DB Clientes carregado: {arquivo_db.name}")

    def carregar_db_cto(self):
        arquivo_db_cto = st.file_uploader(
            "Escolha o arquivo do banco de dados CTO (Excel)",
            type=["xlsx"]
        )
        if arquivo_db_cto:
            self.arquivo_db_cto = arquivo_db_cto
            st.write(f"Arquivo DB CTO carregado: {arquivo_db_cto.name}")

    def processar(self):
        if self.arquivo_kml and self.arquivo_db and self.arquivo_db_cto:
            if self.arquivo_kml.name.lower().endswith('.kmz'):
                kml_files = self.extrair_marcadores()
                if not kml_files:
                    st.warning("Nenhum arquivo KML encontrado no KMZ.")
                else:
                    for kml_file in kml_files:
                        with ZipFile(self.arquivo_kml, 'r') as zip_ref:
                            kml_content = zip_ref.read(kml_file).decode('utf-8')
                            dados_kml = self.extrair_dados_kml(kml_content)
                            dados_cto = self.extrair_dados_cto()
                            self.atualizar_descricao_com_db(dados_kml, dados_cto)
                            self.gerar_kml_csv_atualizado(kml_file, dados_kml)
                            link = self.gerar_link_google_maps(dados_kml[0][1], dados_kml[0][2])
                            st.markdown(f"[Abrir no Google Maps]({link})", unsafe_allow_html=True)
            elif self.arquivo_kml.name.lower().endswith('.kml'):
                kml_content = self.arquivo_kml.read().decode('utf-8')
                dados_kml = self.extrair_dados_kml(kml_content)
                dados_cto = self.extrair_dados_cto()
                self.atualizar_descricao_com_db(dados_kml, dados_cto)
                self.gerar_kml_csv_atualizado(self.arquivo_kml.name, dados_kml)
            else:
                st.warning("Formato de arquivo não suportado. Por favor, selecione um arquivo KML ou KMZ.")
            st.success("Extração concluída")

    def extrair_dados_cto(self):
        wb_cto = openpyxl.load_workbook(self.arquivo_db_cto)
        sheet_cto = wb_cto.active
        dados_cto = {}
        for excel_row in sheet_cto.iter_rows(min_row=2, max_row=sheet_cto.max_row, values_only=True):
            nome_cto_pon_parts = excel_row[0].split(' ', 1)
            nome_sigla = excel_row[0].split('-', 1)[0].strip()
            if len(nome_cto_pon_parts) == 2:
                nome_cto, nome_pon = nome_cto_pon_parts
            else:
                nome_cto = excel_row[0]
                nome_pon = ""
            description_cto = excel_row[3]
            total_portas_utilizadas = excel_row[4] if excel_row[4] is not None else 0
            total_portas_disponiveis = excel_row[5] if excel_row[5] is not None else 0
            dados_cto[nome_cto] = {
                'description_splitter': description_cto,
                'total_portas_utilizadas': total_portas_utilizadas,
                'total_portas_disponiveis': total_portas_disponiveis,
                'nome_cto': nome_cto,
                'nome_pon': nome_pon,
                'nome_sigla': nome_sigla
            }
        return dados_cto

    def calcular_soma_portas_por_nome_pon(self, dados_cto):
        soma_portas = defaultdict(int)
        for cto_info in dados_cto.values():
            nome_pon = cto_info.get('nome_pon')
            nome_sigla = cto_info.get('nome_sigla')
            portas_utilizadas = cto_info.get('total_portas_utilizadas', 0)
            if nome_pon and nome_sigla:
                key = (nome_pon, nome_sigla)
                soma_portas[key] += portas_utilizadas
        return soma_portas

    def extrair_marcadores(self):
        kml_files = []
        with ZipFile(self.arquivo_kml, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                if file_info.filename.lower().endswith('.kml'):
                    kml_files.append(file_info.filename)
        return kml_files

    def extrair_dados_kml(self, kml_content):
        root = ET.fromstring(kml_content)
        dados = []
        for placemark in root.findall(".//{http://www.opengis.net/kml/2.2}Placemark"):
            nome_element = placemark.find(".//{http://www.opengis.net/kml/2.2}name")
            coordenadas_element = placemark.find(".//{http://www.opengis.net/kml/2.2}coordinates")
            description_element = placemark.find(".//{http://www.opengis.net/kml/2.2}description")
            if nome_element is not None and coordenadas_element is not None and description_element is not None:
                nome = nome_element.text
                coordenadas = coordenadas_element.text
                description = description_element.text
                try:
                    longitude, latitude, _ = map(str, coordenadas.split(','))
                    nome_parte = nome.split(' ')[0]
                    nome_sigla = nome.split('-', 1)[0].strip()
                    nome_pon = self.obter_nome_pon_a_partir_do_nome_cto(nome_parte)
                    dados.append([nome_parte, latitude, longitude, description, nome_pon, nome_sigla])
                except ValueError as e:
                    st.error(f"Erro ao converter coordenadas para '{nome}': {coordenadas}. Erro: {e}")
            else:
                st.warning("Elemento ausente em um placemark. Ignorando.")
        return dados

    def obter_nome_pon_a_partir_do_nome_cto(self, nome_cto):
        partes_nome_cto = nome_cto.split(' ')
        if len(partes_nome_cto) > 1:
            nome_pon = partes_nome_cto[1]
            return nome_pon
        else:
            return None

    def atualizar_descricao_com_db(self, dados_kml, dados_cto):
        for row in dados_kml:
            nome = row[0]
            nome_tratado = nome.split(' ')[0]
            nova_descricao = ""
            descricao_cto_info = dados_cto.get(nome_tratado, {})
            if 'description_splitter' in descricao_cto_info:
                nova_descricao += f"Splitter de Atendimento 1x{descricao_cto_info['description_splitter']}\n\n"
                nova_descricao += f"Portas da CTO Utilizadas: {descricao_cto_info.get('total_portas_utilizadas', 'N/A')}\n"
                nova_descricao += f"Portas da CTO Disponíveis: {descricao_cto_info.get('total_portas_disponiveis', 'N/A')}\n\n"
                info_db = self.dados_db.get(nome_tratado, [])
                if info_db:
                    nova_descricao += "Listagem de IDs com Serviço:\n"
                    for entry in info_db:
                        id_formatado = entry['id'].ljust(9)
                        nova_descricao += f" - ID: {id_formatado}, Status: {entry['status']}\n"
                else:
                    nova_descricao += "Nenhum Id Ativo\n"
                nome_sigla = descricao_cto_info.get('nome_sigla', '')
                # Corrected method call here: removed the second argument
                soma_portas = self.calcular_soma_portas_por_nome_pon(dados_cto).get(
                    (descricao_cto_info.get('nome_pon', ''), nome_sigla), 'N/A')
                nova_descricao += f"\nTotal de Clientes na PON: {soma_portas}\n"
            else:
                nova_descricao += "Informações CTO não disponíveis.\n"
            row[3] = self.formatar_descricao_vertical(nova_descricao)

    def formatar_descricao_vertical(self, descricao):
        return "<pre>" + descricao.rstrip('\n') + "</pre>"

    def gerar_kml_csv_atualizado(self, kml_file, dados):
        novo_kml_file = f'{os.path.splitext(kml_file)[0]}_atualizado.kml'
        with open(novo_kml_file, 'w', encoding='utf-8') as f:
            f.write(self.substituir_description_kml(kml_file, dados))
        st.write(f'Arquivo KML atualizado salvo como {novo_kml_file}')

        csv_file = f'{os.path.splitext(kml_file)[0]}_atualizado.csv'
        with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(['Nome', 'Latitude', 'Longitude', 'CTO', 'IDs'])
            for row in dados:
                nome_tratado = row[0]
                latitude = row[1]
                longitude = row[2]
                cto_ids = row[3].replace('\n', '_')
                csv_writer.writerow([nome_tratado, latitude, longitude, cto_ids])
        st.write(f'Dados extraídos e descrição atualizada para {csv_file}')

    def gerar_link_google_maps(self, latitude, longitude):
        return f'https://www.google.com/maps/place/{latitude},{longitude}'

def main():
    st.title("ATUALIZADOR DE MAPAS KML")
    app = KMLProcessorApp()
    app.carregar_kml()
    app.carregar_db()
    app.carregar_db_cto()
    if st.button("Processar"):
        app.processar()

if __name__ == "__main__":
    main()
