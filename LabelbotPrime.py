import os
import tkinter as tk
from tkinter import ttk, filedialog
import win32print
import win32api
import subprocess
import datetime
import re
import time
import win32gui

class EtiquetaGenerator(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.selected_label = tk.StringVar()
        self.master = master
        self.init_ui()

    def init_ui(self):
        # Criar campos de seleção de arquivo
        self.file_label = tk.Label(self, text="Selecionar arquivo XML:")
        self.file_label.pack(pady=10)
        self.file_entry = tk.Entry(self, width=50)
        self.file_entry.pack(pady=5)
        self.file_button = tk.Button(self, text="Pesquisar", command=self.select_file)
        self.file_button.pack(pady=10)

        # Criar campos de entrada para numPed e numNF
        self.num_ped_label = tk.Label(self, text="Número do Pedido:")
        self.num_ped_label.pack(pady=10)
        self.num_ped_entry = tk.Entry(self, width=20)
        self.num_ped_entry.pack(pady=5)

        self.num_nf_label = tk.Label(self, text="Número da NF:")
        self.num_nf_label.pack(pady=10)
        self.num_nf_entry = tk.Entry(self, width=20)
        self.num_nf_entry.pack(pady=5)

        # Criar botão para gerar etiquetas
        self.generate_button = tk.Button(self, text="Gerar Etiquetas", command=self.generate_etiquetas)
        self.generate_button.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, file_path)

    def generate_etiquetas(self):
        file_path = self.file_entry.get()
        num_ped = self.num_ped_entry.get()
        num_nf = self.num_nf_entry.get()

        # Parse XML file
        tree = ET.parse(file_path)
        root = tree.getroot()

        #Horario
        horarioAtual = datetime.now().time()
        horarioFormatado =  horarioAtual.strftime("%H:%M:%S")

        #Data
        data = datetime.now()
        dataFormatada = data.strftime("%d/%m/%Y")

        # Extract data from XML file
        loja_data_list = []
        for pedido in root.findall('./PedidoRiac'):
            for loja in pedido.findall('.//Loja'):
                loja_data = {
                    'CNPJ': loja.find('CNPJ').text if loja.find('CNPJ') is not None else None,
                    'ItemPedido': loja.find('ItemPedido').text if loja.find('ItemPedido') is not None else None,
                    "DescricaoMaterial": loja.find('DescricaoMaterial').text if loja.find('DescricaoMaterial') is not None else None,
                    "QuantidadePecGrade": loja.find('QuantidadePecGrade').text if loja.find('QuantidadePecGrade') is not None else None,
                    "QuantidadePeca": loja.find('QuantidadePeca').text if loja.find('QuantidadePeca') is not None else None,
                    "MaterialForn": loja.find('MaterialForn').text if loja.find('MaterialForn') is not None else None,
                    "NomeForn": loja.find('NomeForn').text if loja.find('NomeForn') is not None else None,
                    "CodigoBareau": loja.find('CodigoBureau').text if loja.find('CodigoBureau') is not None else None,
                    "SemanaEntr": loja.find('SemanaEntr').text if loja.find('SemanaEntr') is not None else None,
                    "SequenciaInicial": loja.find('SequenciaInicial').text if loja.find('SequenciaInicial') is not None else None,
                    "SequenciaFinal": loja.find('SequenciaFinal').text if loja.find('SequenciaFinal') is not None else None,
                    "QuantidadeEtiqueta": loja.find('QuantidadeEtiqueta').text if loja.find('QuantidadeEtiqueta') is not None else None,
                    "centroentrega": loja.find('centroentrega').text if loja.find('centroentrega') is not None else None,
                    "centrofaturamento": loja.find('centrofaturamento').text if loja.find('centrofaturamento') is not None else None
                }

                for volume in loja.findall('.//DADOS_VOLUME/ETIQUETA'):
                    volume_data = {
                        "numVolume": volume.find('numVolume').text if volume.find('numVolume') is not None else None,
                        "itemQuebra": volume.find('itemQuebra').text if volume.find('itemQuebra') is not None else None,
                        "codMercadoria": volume.find('codMercadoria').text if volume.find('codMercadoria') is not None else None,
                        "codBarraCdAtual": volume.find('codBarraCdAtual').text if volume.find('codBarraCdAtual') is not None else None,
                        "footerPagina": volume.find('footerPagina').text if volume.find('footerPagina') is not None else None,
                        "remNome": volume.find('remNome').text if volume.find('remNome') is not None else None,
                        "fluxoDeposito": volume.find('fluxoDeposito').text if volume.find('fluxoDeposito') is not None else None,
                        "qtdeTotal": volume.find('qtdeTotal').text if volume.find('qtdeTotal') is not None else None,
                        "desCodigoDestino": volume.find('desCodigoDestino').text if volume.find('desCodigoDestino') is not None else None,
                        "codDco": volume.find('codDco').text if volume.find('codDco') is not None else None,
                        "destNome": volume.find('destNome').text if volume.find('destNome') is not None else None,
                        "desAuxiliar_1": volume.find('desAuxiliar_1').text if volume.find('desAuxiliar_1') is not None else None,
                        "codBarra": volume.find('codBarra').text if volume.find('codBarra') is not None else None,
                        "desAuxiliar_2": volume.find('desAuxiliar_2').text if volume.find('desAuxiliar_2') is not None else None,
                        "detalhesCodBarra": volume.find('detalhesCodBarra').text if volume.find('detalhesCodBarra') is not None else None,
                        "desEndereco": volume.find('desEndereco').text if volume.find('desEndereco') is not None else None,
                        "tipoProduto": volume.find('tipoProduto').text if volume.find('tipoProduto') is not None else None,
                        "numAuxiliar": volume.find('numAuxiliar').text if volume.find('numAuxiliar') is not None else None,
                        "codArtigo": volume.find('codArtigo').text if volume.find('codArtigo') is not None else None
                    }

                    loja_data_combined = {**loja_data, **volume_data}
                    loja_data_list.append(loja_data_combined)

        # Convert data to JSON
        json_data = loja_data_list

        # Create output directory if it doesn't exist
        output_dir = "Etiquetas"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        
        # Generate etiquetas
        for Etiqueta in json_data:
            #Acessa os dados dos Itens
            cnpj = Etiqueta.get('CNPJ', '')
            numVolume = Etiqueta.get('numVolume', '')
            itemQuebra = Etiqueta.get('itemQuebra', '')
            codMercadoria = Etiqueta.get('codMercadoria', '')
            codBarraCdAtual = Etiqueta.get('codBarraCdAtual', '')
            footerPagina = Etiqueta.get('footerPagina', '')
            remNome = Etiqueta.get('remNome', '')
            fluxoDeposito = Etiqueta.get('fluxoDeposito', '')
            qtdeTotal = Etiqueta.get('qtdeTotal', '')
            desCodigoDestino = Etiqueta.get('desCodigoDestino', '')
            codDco = Etiqueta.get('codDco', '')
            destNome = Etiqueta.get('destNome', '')
            desAuxiliar_1 = Etiqueta.get('desAuxiliar_1', '')
            codBarra = Etiqueta.get('codBarra', '')
            desAuxiliar_2 = Etiqueta.get('desAuxiliar_2', '')
            detalhesCodBarra = Etiqueta.get('detalhesCodBarra', '')
            desEndereco = Etiqueta.get('desEndereco', '').replace('GalpÃ£o', 'Galpão')
            tipoProduto = Etiqueta.get('tipoProduto', '')
            numAuxiliar = Etiqueta.get('numAuxiliar', '')
            codArtigo = Etiqueta.get('codArtigo', '')

            tamanhodesAuxiliar_1 = Etiqueta.get('desAuxiliar_1')
            IdentVolume = numVolume.replace("/","-")
            cep = tamanhodesAuxiliar_1[:14]
            cidadeEstado =desAuxiliar_1[14:]

            codigo_zpl_atualizado = f"""
^XA
^CI28
^MMT
^PW799
^LL0759
^LS0
^FO0,128^GFA,06144,06144,00012,:Z64:
eJztmL9q3EAQxkfagAI2LFelOaLDVXBCzuVBhPRaKQK34Oa6PEKuDJfCbcCBk58gr6AyuEiuCRzGaKNG830yETnlJPsMnuqH0O5+s/9mZkWe7MDsOXECDG/BEfEabDx46rfKc6945MGx/6WclRa8AacOnMjfv/P/3E/sv5M0jMt6xPpCOSD9krX4K8w0P415I3PAFxvwuqRhfU46VU/kvTawNxc6n3EeKL8UYX5fcyoyI5FJ35yenChnHvPWldtsQpwMwJV+HWIf/clisRhS5442xNC982g0+u898+7y8qs8bntF/OBr0canq9W3moc4112t67hvy5+dNAw9sXvdLQPoeSPdbCJyXPOZyLjmWSMuBNDvDDhH3MyKTOOU3c6Lmo9K72oOOEZ7xEeZI5WQjHjOqQfaRreIp3ab6VjZJtZ4mi2tdpQJNKcU3ysfP9R81pzGY+lgz3iuQgeOOrFp4Zg4JU6IZy3f05Z+THqh90DwFN+H0/k4LVYKc5yvKJ8jX80pd71CLRBfbcHBBvmqKTRfHZsC+585yolbvltmp+c3nTqM65E/W39OHKlmUxr1K9iaXB1eGic9GOczu+U/VGeJJc5+gHGvSoh7VSK6Sy3fpVSPYBmrexifTVvTm37OV2cLHThuraf0pzv1lHJVT2njqp4Cu3/XU+P9aijt507/2J9C+3NpwL/hMPtiqH4MyfdGnfuaFm834xincb9So2sXlYj7tkQMjT3Fff8RftlPqj8162XNiYHOhDZcAtcr3hAXxHknTr9cf4a2+8xvJ0Det750NZqWdwDL7wDX9A7gAsphGvsW917LO0Da1ztAT/XCpDLpZvp/T/nVvTLlVwdRZ3WNRw0NkZOH5tPpSuvrQ5jPJ9vN/gAbPRsV:584D
^FO64,64^GFA,07296,07296,00012,:Z64:
eJzdmL9PGzEUx31nVKMg3bUbEilBTBVTxyugkEr9Mzr0T6ATaYXIVR0Y2TuUCJYqSJ0jWqWROnTtn3CQgQGpTcVASlFSu7HvfYPO3IUk/fWkSB8p9vPX9jv7+TH2b5rLFmLm7EnMImwTs/WYPRbcmNEPP41idnpNEnSYQTQvEe6QH7FLmv3XoCEPHKSzKJNO0Saf/IjGYmGYQegI5oJ/PgK74xSVYIc0Fu4j7q9gtIbjih8Rkn9+DtyrkrZKmni7+b1OIueKb6hRnt2Yrf5r4D+DYXubT+bQ+jMeERfAUTe5+cxxNdW/c+f2MJKtfgY0W8ytQjy0y9S3Q33FHMTbXJUY5uvNA+ct/LiUqtPdI/8DNg0McYjxKc7hvApBP7tPGob9XhaJh11btBz6zE+AG/Wx6HSrxOKS2G8A4157Q/LbDN/Xb7T/Kf7x3J5IvgE+7flGmCjfG5Zt5/C7enL7DDyVqCyb+b0ImNYht4RxG2RiSiHWZcyR/7UucaWn/dcYiyD/SGQZD+uijyoenn/QfE4+1d7l9JJeFw/iM7Hpq9joUXysx1U+jQYVD4ZVPBg9Kv80epStwb1o9AzGcxTzr/N8U6+JPM8DMxd5nj8DPa39kPhlMj/9pNsv1tmK7qviqtAlzmVIFQTc5TgXsz7KjtP2i9N+qdwyXh/JsU/JsZ6rueg9EJQt5JI5yWRwLi1plvtS1Huj9qX3ta9DyPj/9rHPam2PLNxq6fYy3irajzhos2KRfE6Z87EGc7xiqGfLoufMjOva9ZyOQ4/6eJdJz+oW9d0CPRvbjNYB4vAI2MSD0lMWpCfwyGd8aFn1hKRH+jN6FHe1HsUb28RGj2KjRxnGajmOzyrp2SU9fgPO5ADyK/nuc/Ud4M0G8rtq9tu86rDGQb+PWp82W43nWDWX2zXvi6aZrryvI82iHTCTovjvA/ZFy/fnA3aiv1PPl+t/ofWIOosqmnmHRbk+C6cT++TynKlqdtgsKJgevIPhXq81wph7vT6rd4RhdbZs6vaKi5qF/FvkzbjypyfjqMmaiSlhqzCu5RJr6j/U2l+4/TuSy7Wv6LtKyHX2da4g8uvM0eeq58H5b2GxCff+RRSz0y4liwHDfIafQH7ehfwqD/Fjy2+LoAfyfN4iPexhM1XPSDZS3QB40iZKJMFSz/HqGeo2eagPwL3sANvyvWHNlg/zvQUaF4dqWniBpZoT0Rx5B+Z7afkWBDCH/NMBzfjeDOkd4YZpF+6kbBn4AfBCmNrVVg/cSa4H3ppEPfBFusypSb+pga31paUMSeN1+a3xOa56xd546gDM8i4WF+BzLqJxCyEwucnyLvbvUr1rfHU80mN7uzF4Bw2dPIeoE/YRxpp6lEXpKAl8MosNqGP8sRoLJkug0wUNBdgLnSspE8D+PjDcpz7kreKM2IW+A/5ncO5lYHxMJZsD5VhrHaA2fB2AjOJw5kdyLOWs7yNY5xLlqzyC8+Q7sbdCPr1Ck3g+BCaX1lpW4++q443bfgJU4hRJ:F103
^FWB                 
^FT23,20            
^FO160,0             
^GB3,700,3^FS      
^FT230,690^A0B,40,36^FH\^FDCliente: LOJAS RIACHUELO SA 1^FS
^FT295,690^A0B,40,36^FH\^FDRua: {desEndereco}^FS
^FT360,690^A0B,40,36^FH\^FDBairro: ITAIPAVA^FS
^FT360,310^A0B,40,36^FH\^FDDISTRIBUIÇÃO^FS
^FT430,690^A0B,40,36^FH\^FD{cep}^FS
^FT430,240^A0B,40,36^FH\^FD{destNome}^FS
^FT500,690^A0B,40,36^FH\^FDCidade: {cidadeEstado}^FS
^FT500,210^A0B,40,36^FH\^FDC180^FS
^FT570,690^A0B,40,36^FH\^FDNúm Pedido: {num_ped}^FS
^FT640,690^A0B,40,36^FH\^FDNota Fiscal: {num_nf}^FS
^FT640,340^A0B,40,36^FH\^FDVolume: {numVolume} ^FS
^FT700,690^A0B,40,36^FH\^FDTrasnportadora: HEDRONS TÊXTIL LTDA^FS
^FT743,690^A0B,28,28^FH\^FD{dataFormatada}^FS
^FT743,560^A0B,28,28^FH\^FD{horarioFormatado}^FS
^FO00,700^GB1000,3,3^FS
^FT40,1320^A0B,40,36^FH\^FDHEDRONS TEXTIL LTDA^FS
^FT75,1320^A0B,30,26^FH\^FDDestinatario^FS
^FT125,1320^A0B,36,30^FH\^FDLOJAS RIACHUELO S.A^FS
^FT165,1320^A0B,30,26^FH\^FD{desEndereco}^FS
^FT205,1320^A0B,30,26^FH\^FD{desAuxiliar_1}C^FS
^FT255,1320^A0B,30,26^FH\^FD{desAuxiliar_2}^FS
^FWB
^BY3,15,60
^FT360,1320^BC,60,N^FD12345678^FS
^FT420,1320^A0B,38,32^FH\^FDCodigo Artigo^FS
^FT420,950^A0B,38,32^FH\^FDDestino^FS
^FT470,1320^A0B,38,32^FH\^FD{codBarraCdAtual}^FS
^FT470,950^A0B,38,32^FH\^FD{desCodigoDestino}^FS
^FT520,1320^A0B,38,32^FH\^FDPedido^FS
^FT520,950^A0B,38,32^FH\^FDQuant. Total^FS
^FT560,1320^A0B,38,32^FH\^FD{numAuxiliar}^FS
^FT560,950^A0B,38,32^FH\^FD3^FS
^FT600,1320^A0B,38,32^FH\^FDVolumes^FS
^FT600,1100^A0B,38,32^FH\^FDDCO^FS
^FT600,950^A0B,38,32^FH\^FDTipo Produto^FS
^FT640,1320^A0B,38,32^FH\^FD400^FS
^FT640,950^A0B,38,32^FH\^FDN. {numVolume}^FS
^FT165,1320^A0B,30,26^FH\^FD{desEndereco}^FS
^FT205,1320^A0B,30,26^FH\^FD{desAuxiliar_1}^FS
^FWB
^BY3,15,60
^FT740,1320^BC^FD{codBarra}^FS
^PQ1,0,1,Y^XZ
"""
        # Create output file name
            output_file_name = f"{IdentVolume} - {codMercadoria} - {codBarraCdAtual}.zpl"

            # Create output file path
            output_file_path = os.path.join(output_dir, output_file_name)

            # Write ZPL code to output file
            with open(output_file_path, 'w', encoding='utf-8') as file:
                file.write(codigo_zpl_atualizado)

            print(f"File {output_file_name} saved successfully!")

class ExportTab(ttk.Frame): 
    print("teste")


class PrintTab(ttk.Frame): 
    def __init__(self, master):
        super().__init__(master)
        self.selected_label = tk.StringVar()
        self.master = master
        self.file_list = []  # Lista de nomes de arquivos
        self.sort_order = True  # True para crescente
        self.init_ui()

    def init_ui(self):
        self.folder_path = tk.StringVar()
        self.selected_printer = tk.StringVar()

        # Botão de seleção de pasta
        self.folder_button = ttk.Button(self, text="Selecionar Pasta", command=self.select_folder)
        self.folder_button.pack(pady=10)

        # Mostrar pasta selecionada
        self.folder_label = ttk.Label(self, textvariable=self.folder_path)
        self.folder_label.pack()

        # Campo de entrada para filtro
        self.filter_label = tk.Label(self, text="Filtrar por nome do arquivo:")
        self.filter_label.pack(pady=5)

        self.filter_entry = tk.Entry(self, width=50)
        self.filter_entry.pack(pady=5)
        self.filter_entry.bind("<KeyRelease>", lambda event: self.update_treeview())  # Atualiza a visualização ao digitar

        # Treeview para arquivos
        self.tree = ttk.Treeview(self, columns=("File Name", "File Path"), show="headings")
        #self.tree.column("File Name", text="Id")
        self.tree.heading("File Name", text="Nome do Arquivo")
        self.tree.heading("File Path", text="Caminho Completo")
        self.tree.pack(pady=10)

        # Combobox para etiquetas
        self.label_combobox = ttk.Label(self, text="Selecionar Etiqueta:")
        self.label_combobox.pack(pady=5)

        self.label_selector = ttk.Combobox(self, textvariable=self.selected_label, values=["ET0087", "ET0240"])
        self.label_selector.pack(pady=5)
        self.label_selector.bind("<<ComboboxSelected>>", lambda event: self.update_label_size())

        # Combobox para impressora
        self.printer_label = ttk.Label(self, text="Selecionar Impressora:")
        self.printer_label.pack(pady=5)

        self.printer_combobox = ttk.Combobox(self, textvariable=self.selected_printer)
        self.printer_combobox.pack(pady=5)
        self.load_printers()

        # Botão para imprimir arquivos
        self.print_button = ttk.Button(self, text="Imprimir Arquivos Selecionados", command=self.imprimir_selecionados)
        self.print_button.pack(pady=10)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path.set(folder_path)  # Atualiza o caminho da pasta
            # Obter a lista de arquivos da pasta
            self.file_list = os.listdir(folder_path)
            # Ordenar a lista de arquivos (crescente) com base nos números
            self.file_list.sort(key=self.natural_sort_key)  
            self.update_treeview()  # Atualizar a visualização


    def natural_sort_key(self, text):
    # Função para separar números de texto e ordenar corretamente
        return [int(part) if part.isdigit() else part.lower() for part in re.split(r'(\d+)', text)]

    def update_treeview(self):
        # Limpar a Treeview antes de adicionar novos arquivos
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Obter o texto do filtro
        filter_text = self.filter_entry.get().lower()

        # Filtrar e ordenar arquivos
        filtered_files = [file_name for file_name in self.file_list if filter_text in file_name.lower()]
        filtered_files.sort(key=self.natural_sort_key)  # Ordenar em ordem crescente

        # Adicionar arquivos à Treeview com base no filtro
        for file_name in filtered_files:
            full_path = os.path.join(self.folder_path.get(), file_name)
            self.tree.insert("", "end", values=(file_name, full_path))

    def sort_files(self):
        # Alternar a ordem de exibição
        self.sort_order = not self.sort_order  
        # Ordenar a lista de arquivos (crescente ou decrescente)
        self.file_list.sort(key=lambda x: x.lower(), reverse=not self.sort_order)  
        self.update_treeview()  # Atualizar Treeview

    def load_printers(self):
        # Enumera as impressoras disponíveis e as adiciona ao combobox
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        self.printer_combobox['values'] = printers

    def update_label_size(self):
        # Função que retorna as dimensões da etiqueta selecionada
        label = self.selected_label.get()
        if label == "ET0087":
            width, height = 7.51, 10  # Largura e altura em cm
        elif label == "ET0240":
            width, height = 10, 17  # Largura e altura em cm
        else:
            width, height = None, None

        if width and height:
            print(f"Etiqueta selecionada: {label} - Tamanho: {width} cm x {height} cm")
        else:
            print("Etiqueta não encontrada.")

    def imprimir_selecionados(self):
        selected_items = self.tree.selection()  # Obtém os itens selecionados
        if selected_items:
            files = [self.tree.item(item, "values")[1] for item in selected_items]  # Caminhos dos arquivos selecionados
            print("Arquivos selecionados para impressão:", files)

            printer = self.selected_printer.get()
            if not printer:
                print("Nenhuma impressora selecionada.")
                return
            
            for file_path in files:
                try:
                    # Normaliza o caminho do arquivo
                    file_path = os.path.normpath(file_path)

                    # Verifica se o arquivo existe
                    if os.path.exists(file_path):
                        # Se for um arquivo PDF, usa o Adobe Reader  
                       # Envia o arquivo para a impressora selecionada
                        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
                        print(f"Arquivo {file_path} enviado para a impressora {printer}.")
                    else:
                        print(f"Arquivo {file_path} não encontrado.")
                except Exception as e:
                    print(f"Erro ao enviar {file_path} para a impressora: {e}")
        else:
            print("Nenhum arquivo foi selecionado.")



if __name__ == "__main__":
    root = tk.Tk()
    root.title("Labelbot Prime")

    # Criar o notebook para as abas
    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    # Adicionar a aba "Gerador de Etiquetas"
    etiqueta_generator = EtiquetaGenerator(notebook)
    notebook.add(etiqueta_generator, text="Gerador de Etiqueta")

    # Adicionar a aba "Imprimir"
    print_tab = PrintTab(notebook)
    notebook.add(print_tab, text="Imprimir")

    root.mainloop()