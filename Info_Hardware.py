import os
import wmi
import psutil
import platform
import threading
import pythoncom  # Necessário para inicializar o COM dentro de threads

# Funções para coleta de informações de hardware
def get_motherboard_model():
    try:
        pythoncom.CoInitialize()  # Inicializa o COM para esta thread
        c = wmi.WMI()
        for board in c.Win32_BaseBoard():
            return board.Product
    except Exception as e:
        print(f"Erro ao obter informações da placa-mãe: {e}")
        return "Informação não disponível"

def get_ram_info():
    try:
        pythoncom.CoInitialize()  # Inicializa o COM para esta thread
        total_memory = round(psutil.virtual_memory().total / (1024 ** 3), 2)
        used_slots = len(wmi.WMI().Win32_PhysicalMemory())
        return total_memory, used_slots
    except Exception as e:
        print(f"Erro ao obter informações de memória: {e}")
        return None, None

def get_disk_info():
    try:
        disk_info = psutil.disk_partitions()
        disk_model = []
        for partition in disk_info:
            if 'fixed' in partition.opts:
                usage = psutil.disk_usage(partition.mountpoint)
                disk_model.append((partition.device, round(usage.total / (1024 ** 3), 2)))
        return disk_model
    except Exception as e:
        print(f"Erro ao obter informações do HD: {e}")
        return []

def get_processor_info():
    try:
        physical_cores = psutil.cpu_count(logical=False)
        frequency = psutil.cpu_freq().current
        return physical_cores, frequency
    except Exception as e:
        print(f"Erro ao obter informações do processador: {e}")
        return None, None

def get_graphics_card_info():
    try:
        pythoncom.CoInitialize()  # Inicializa o COM para esta thread
        c = wmi.WMI()
        for card in c.Win32_VideoController():
            return card.Name
    except Exception as e:
        print(f"Erro ao obter informações da placa de vídeo: {e}")
        return "Informação não disponível"

def get_network_adapter_info():
    try:
        pythoncom.CoInitialize()  # Inicializa o COM para esta thread
        c = wmi.WMI()
        adapters = c.Win32_NetworkAdapter()
        adapter_list = [adapter.Name for adapter in adapters if adapter.NetConnectionStatus == 2]
        return adapter_list
    except Exception as e:
        print(f"Erro ao obter informações da placa de rede: {e}")
        return []

def get_windows_info():
    try:
        version = platform.version()
        build = platform.build()
        return version, build
    except Exception as e:
        print(f"Erro ao obter informações do Windows: {e}")
        return "Informação não disponível", "Informação não disponível"

def get_formatting_date():
    try:
        c = wmi.WMI()
        os = c.Win32_OperatingSystem()[0]
        return os.InstallDate.split('.')[0]  # Data de formatação em formato ISO
    except Exception as e:
        print(f"Erro ao obter a data de formatação do Windows: {e}")
        return "Informação não disponível"

# Função para gerar o relatório em HTML
def generate_report():
    motherboard_model = get_motherboard_model()
    total_memory, used_slots = get_ram_info()
    disks = get_disk_info()
    physical_cores, frequency = get_processor_info()
    graphics_card = get_graphics_card_info()
    network_adapters = get_network_adapter_info()
    windows_version, windows_build = get_windows_info()
    formatting_date = get_formatting_date()

    # Geração do relatório HTML com Bootstrap
    disk_info_list = "".join([f"<li>{disk[0]}: <strong>{disk[1]}GB</strong></li>" for disk in disks])
    network_info_list = "".join([f"<li>{adapter}</li>" for adapter in network_adapters])

    html_content = f"""
    <html>
        <head>
            <title>Relatório de Hardware</title>
            <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
            <style>
                body {{
                    padding: 20px;
                    background-color: #f8f9fa;
                }}
                .list-group-item {{
                    background-color: #ffffff;
                    border: 1px solid #dddddd;
                    margin-bottom: 10px;
                    border-radius: 5px;
                }}
                .list-group-item strong {{
                    color: #007bff;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1 class="text-center text-primary">Relatório de Informações de Hardware</h1>
                <ul class="list-group">
                    <li class="list-group-item">
                        <strong>Memória RAM:</strong> {total_memory if total_memory else 'Informação não disponível'}GB | 
                        <strong>Ocupando {used_slots if used_slots else 'Informação não disponível'} slots</strong>
                    </li>
                    <li class="list-group-item"><strong>Modelo da Placa Mãe:</strong> {motherboard_model}</li>
                    <li class="list-group-item"><strong>Placa de Vídeo:</strong> {graphics_card}</li>
                    <li class="list-group-item"><strong>Processador:</strong> {physical_cores if physical_cores else 'Informação não disponível'} núcleos físicos, Frequência: {frequency if frequency else 'Informação não disponível'}MHz</li>
                    <li class="list-group-item"><strong>Discos:</strong>
                        <ul>{disk_info_list}</ul>
                    </li>
                    <li class="list-group-item"><strong>Adaptadores de Rede:</strong>
                        <ul>{network_info_list}</ul>
                    </li>
                    <li class="list-group-item"><strong>Versão do Windows:</strong> {windows_version}</li>
                    <li class="list-group-item"><strong>Build do Windows:</strong> {windows_build}</li>
                    <li class="list-group-item"><strong>Data de Formatação do Windows:</strong> {formatting_date}</li>
                </ul>
            </div>
        </body>
    </html>
    """
    return html_content

# Função para rodar o relatório em thread separada
def generate_report_thread():
    html_content = generate_report()

    # Salvar o relatório em arquivo HTML
    try:
        with open("relatorio_hardware.html", "w", encoding="utf-8") as file:
            file.write(html_content)
            print("Relatório de hardware gerado com sucesso.")
    except Exception as e:
        print(f"Erro ao salvar o relatório: {e}")

# Executa a geração do relatório em uma thread
if __name__ == "__main__":
    threading.Thread(target=generate_report_thread).start()
