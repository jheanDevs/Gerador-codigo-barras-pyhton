import os
import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
import flet as ft
import time

# Função para gerar o código de barras
def gerar_codigo_barras(matricula, nome_pasta):
    file_path = os.path.join(nome_pasta, f"{matricula}.png")
    
    # Verifica se o código de barras já foi gerado
    if os.path.exists(file_path):
        return False  # Código de barras já existe, então pula
    else:
        # Gera o código de barras
        with open(file_path, 'wb') as f:
            Code128(matricula, writer=ImageWriter()).write(f)
        return True

# Função para processar a planilha
def processar_planilha(arquivo, pasta_destino, page, progress_bar, log_container):
    df = pd.read_excel(arquivo)

    total_linhas = len(df)
    novos_codigos = 0

    for index, row in df.iterrows():
        matricula = str(row['CHAPA'])
        nome_funcionario = row['NOME'].split()[:2]  # Pegando o primeiro e segundo nome
        nome_exibido = " ".join(nome_funcionario)
        
        if gerar_codigo_barras(matricula, pasta_destino):
            novos_codigos += 1
            log_mensagem(f"Código de barras para {nome_exibido} gerado com sucesso!", ft.colors.GREEN, log_container, page)
        else:
            log_mensagem(f"Código de barras para {nome_exibido} já existe. Pulando...", ft.colors.ORANGE, log_container, page)
        
        # Atualização visual da barra de progresso
        progress_bar.value = (index + 1) / total_linhas
        progress_bar.update()
        page.update()

    # Feedback final ao usuário
    if novos_codigos == 0:
        log_mensagem(f"Todos os códigos já existiam. Nenhum novo código gerado.", ft.colors.BLUE, log_container, page)
    else:
        log_mensagem(f"{novos_codigos} códigos novos gerados!", ft.colors.GREEN, log_container, page)

    # Mensagem de conclusão final
    log_mensagem("Processo concluído com sucesso!", ft.colors.GREEN, log_container, page)
    
    # Exibir um diálogo de conclusão com novas cores
    page.dialog = ft.AlertDialog(
        title=ft.Text("Processo Concluído!", color=ft.colors.BLACK, size=18),
        content=ft.Text(f"{novos_codigos} códigos novos foram gerados.", color=ft.colors.GREEN, size=14),
        actions=[
            ft.TextButton(
                "OK",
                on_click=lambda e: fechar_aplicacao(page)
            )
        ],
        modal=True,
        bgcolor=ft.colors.WHITE,  # Cor de fundo alterada
    )
    page.dialog.open = True
    page.update()

# Função para fechar o aplicativo
def fechar_aplicacao(page):
    page.window_close()

# Função para adicionar mensagens no log com rolagem
def log_mensagem(mensagem, cor, container, page):
    container.controls.append(
        ft.Text(mensagem, color=cor, size=14)
    )
    page.update()
    time.sleep(0.1)  # Tempo de exibição mais suave
    if len(container.controls) > 6:  # Limita a quantidade de mensagens visíveis
        container.controls.pop(0)
    page.update()

# Função principal do app
def main(page: ft.Page):
    page.title = "Gerar Códigos de Barras"
    page.window_width = 600
    page.window_height = 500
    page.bgcolor = ft.colors.BLACK

    # Barra de progresso estilizada
    progress_bar = ft.ProgressBar(
        width=300, value=0, color=ft.colors.GREEN,
        bgcolor=ft.colors.BLUE_900
    )

    # Contêiner para mensagens de log
    log_container = ft.Column(scroll="always", expand=True)

    # Função chamada após o upload
    def upload_planilha(e: ft.FilePickerResultEvent):
        if e.files:
            planilha_path = e.files[0].path
            output_folder = "C:/Users/Funcionario/Desktop/codigos_barras"
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            processar_planilha(planilha_path, output_folder, page, progress_bar, log_container)
        else:
            log_mensagem("Nenhum arquivo foi selecionado.", ft.colors.RED, log_container, page)

    # Interface de upload
    file_picker = ft.FilePicker(on_result=upload_planilha)
    page.overlay.append(file_picker)

    # Layout da página com melhorias
    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.ElevatedButton(
                        "Selecione a Planilha",
                        on_click=lambda _: file_picker.pick_files(
                            allow_multiple=False,
                            allowed_extensions=["xlsx"]
                        ),
                        style=ft.ButtonStyle(
                            bgcolor=ft.colors.BLUE,
                            color=ft.colors.WHITE,
                            shape=ft.RoundedRectangleBorder(radius=10),
                            overlay_color=ft.colors.BLUE_700,
                        ),
                    ),
                    ft.Text("Progresso:", size=14, color=ft.colors.WHITE),
                    progress_bar,
                    log_container,
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=20,
            alignment=ft.alignment.center,
        )
    )

if __name__ == "__main__":
    ft.app(target=main)
