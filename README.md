# Inventário de Rede ICAVC

Este é um aplicativo em Python para escanear redes locais e gerar relatórios em Excel com informações dos computadores detectados na rede.

## Funcionalidades

- Varre sub-redes configuradas (ex: 10.1.0.0/24)
- Detecta IPs ativos
- Resolve nomes de host
- Tenta identificar o sistema operacional
- Exporta dados em planilha Excel
- Interface gráfica com barra de progresso
- Modo agressivo para escaneamento mais profundo
- Ícone personalizado

##  Requisitos

- Python 3.10+
- Nmap instalado (e no PATH)
- Bibliotecas:
  - `tkinter`
  - `openpyxl`

## Como usar (via Python)

```bash
pip install -r requirements.txt
python inventario.py
