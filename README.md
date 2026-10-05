# Validador Pós-Migração · Agante Tecnologia | Senior HCM

Compara o **Layout Inicial**, a **extração do Senior** e o **Layout Ajustado** campo a campo. Mostra o que foi enviado, o que ficou no Senior, o que precisou ser ajustado e a situação final de cada registro.

## Rodar no computador (Windows)

1. Instale o Python 3.10 ou superior (python.org). Na instalação, marque **Add Python to PATH**.
2. Descompacte esta pasta e dê dois cliques em **`iniciar_validador.bat`**.
   - Na primeira vez, ele instala as dependências.
   - Depois, abre o validador no navegador em `http://localhost:8501`.

Pelo terminal, o equivalente é:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Para outras pessoas da rede acessarem, use `streamlit run app.py --server.address 0.0.0.0`. Elas abrem `http://<IP-da-máquina>:8501`.

## Publicar no Streamlit Community Cloud (mesmo modelo do validador de pré-migração)

1. Crie um repositório no GitHub e envie o conteúdo desta pasta: `app.py`, `engine.py`, `requirements.txt`, `.streamlit/`, `assets/` e `configs/`.
2. Acesse share.streamlit.io e entre com a conta do GitHub. Clique em **Create app**.
3. Escolha o repositório, a branch `main` e o arquivo `app.py`.
4. Defina o endereço (ex.: `validador-pos-migracao-agante`) e clique em **Deploy**.

Para restringir o acesso, use as opções de compartilhamento do app no Streamlit Cloud ou mantenha o repositório privado.

## Formas de carregar os arquivos (por etapa)

| Modo | Como usar | Onde funciona |
|---|---|---|
| **Upload** | Arraste um ou vários arquivos XLSX, XLS, CSV ou TXT | Sempre |
| **Leitura (colar)** | Cole o conteúdo do CSV ou as células copiadas do Excel, com o cabeçalho | Sempre |
| **Caminho na máquina** | Cole o caminho do arquivo ou da pasta, um por linha (ex.: `C:\Migracao\Cliente`). Uma pasta é lida inteira | Quando o validador roda no próprio computador ou tem acesso à pasta de rede. **Não funciona no Streamlit Cloud**, que não enxerga os arquivos da sua máquina |

Em todos os modos, a leitura é automática:

- O cabeçalho é encontrado sozinho, mesmo com títulos acima.
- O separador do CSV/TXT é identificado.
- Todas as abas de uma planilha são lidas.
- Vários arquivos na mesma etapa são somados.

## Logos

Coloque os arquivos `assets/logo_agante.png` e `assets/logo_senior.png`. Sem eles, o cabeçalho mostra os nomes em texto.

## Layouts salvos

As configurações (chave, DE/PARA, tipos, tolerâncias e regras) ficam em `configs/layouts.json`.

No Streamlit Cloud esse arquivo é reiniciado a cada nova publicação. Por isso, use **Exportar/Importar layouts (JSON)** na tela Configuração de Layouts, ou faça commit do `layouts.json` no repositório.

## Privacidade

Os arquivos são processados apenas na sessão de quem está validando. Nenhum dado de colaborador é gravado.
