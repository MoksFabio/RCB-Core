# Guia de Implantação no PythonAnywhere

Este guia passo a passo ajudará você a colocar seu projeto Django "RCB" no ar usando o PythonAnywhere.

## 1. Preparação (Local)

Eu já configurei seu `settings.py` para funcionar tanto no seu computador quanto no PythonAnywhere.

Para facilitar o envio dos arquivos, criei um script que compacta seu projeto ignorando arquivos desnecessários (como `.git` e cache).

1.  Abra o terminal na pasta do projeto.
2.  Execute o script de preparação (o arquivo Python que criei para você):
    ```bash
    python prepare_deploy.py
    ```
    Isso criará um arquivo chamado `rcb_deploy.zip`.

## 2. Upload para o PythonAnywhere

1.  Crie uma conta em [www.pythonanywhere.com](https://www.pythonanywhere.com/).
2.  Vá para a aba **Files**.
3.  Clique no botão **"Upload a file"** e selecione o arquivo `rcb_deploy.zip` que criamos.
4.  Após o upload, abra um **Bash console** (na aba "Consoles").
5.  No console, descompacte o arquivo com o comando:
    ```bash
    unzip -o rcb_deploy.zip -d mysite
    ```
    *(Isso criará uma pasta chamada `mysite` com seu projeto dentro).*

## 3. Configuração do Ambiente Virtual

Ainda no **Bash console** do PythonAnywhere:

1.  Entre na pasta do projeto:
    ```bash
    cd mysite
    ```
2.  Crie um ambiente virtual:
    ```bash
    python3.10 -m venv venv
    ```
    *(Verifique qual versão do Python você quer usar. O padrão costuma ser 3.9 ou 3.10).*
3.  Ative o ambiente virtual:
    ```bash
    source venv/bin/activate
    ```
4.  Instale os requisitos:
    ```bash
    pip install -r requirements.txt
    ```

## 4. Configuração do Web App

1.  Vá para a aba **Web** no painel do PythonAnywhere.
2.  Clique em **"Add a new web app"**.
3.  Escolha **Manual configuration** (a opção "Django" cria um projeto novo, e nós queremos usar o existente).
4.  Escolha a versão do Python correspondente à que você usou no passo 3 (ex: Python 3.10).
5.  **Virtualenv**:
    *   Vá até a seção "Virtualenv".
    *   Digite o caminho para o seu ambiente virtual: `/home/SEU_USUARIO/mysite/venv` (substitua SEU_USUARIO pelo seu nome de usuário no PythonAnywhere).

## 5. Configuração do Código (WSGI)

1.  Na aba **Web**, seção "Code":
    *   **Source code**: `/home/SEU_USUARIO/mysite`
    *   **Working directory**: `/home/SEU_USUARIO/mysite`
2.  **WSGI configuration file**:
    *   Clique no link do arquivo WSGI (algo como `/var/www/seu_usuario_pythonanywhere_com_wsgi.py`).
    *   Apague tudo o que estiver lá e substitua por este conteúdo (ajustando o nome da pasta e do usuário):

    ```python
    import os
    import sys

    # Ajuste o caminho para a pasta do seu projeto
    path = '/home/SEU_USUARIO/mysite'
    if path not in sys.path:
        sys.path.append(path)

    # Diga ao Django qual arquivo de settings usar
    os.environ['DJANGO_SETTINGS_MODULE'] = 'RCB.settings'

    # Ative a aplicação
    from django.core.wsgi import get_wsgi_application
    application = get_wsgi_application()
    ```
    *   Clique em **Save**.

## 6. Arquivos Estáticos

1.  Na aba **Web**, vá para a seção **Static files**.
2.  Adicione um novo mapeamento:
    *   **URL**: `/static/`
    *   **Directory**: `/home/SEU_USUARIO/mysite/staticfiles`
    *(Nota: No seu settings.py, o STATIC_ROOT aponta para 'staticfiles', então usaremos essa pasta).*
3.  Volte para o **Bash console** e rode o comando para coletar os estáticos:
    ```bash
    python manage.py collectstatic
    ```

## 7. Finalização

1.  Volte para a aba **Web**.
2.  Clique no botão verde **Reload**.
3.  Acesse o link do seu site (algo como `seu_usuario.pythonanywhere.com`).

---

## 8. Como Atualizar o Projeto (Futuro)

Se você mudar algo no código do seu computador e quiser enviar para o site (Deploy), **não precisa refazer tudo**. Siga estes passos:

1.  **Gere o novo zip:** Rode no seu computador: `python prepare_deploy.py`.
2.  **Upload:** Na aba **Files** do PythonAnywhere, faça upload do arquivo `rcb_deploy.zip` novamente (pode substituir o antigo).
3.  **Atualize os arquivos:** Abra o **Bash console** e rode:
    ```bash
    unzip -o rcb_deploy.zip -d mysite
    ```
    *(O parâmetro `-o` serve para sobrescrever os arquivos antigos pelos novos).*
4.  **Atualize o Banco (Opcional):** Se você criou novas tabelas no banco de dados, rode:
    ```bash
    workon venv  (ou 'source venv/bin/activate' se o workon não funcionar)
    cd mysite
    python manage.py migrate
    ```
5.  **Atualize os Estáticos (Opcional):** Se mudou CSS ou imagens, rode:
    ```bash
    python manage.py collectstatic
    ```
6.  **Recarregue:** Vá na aba **Web** e clique no botão verde **Reload**.

### ⚠️ PERIGO: CUIDADO COM O BANCO DE DADOS
O arquivo `rcb_deploy.zip` **contém o seu banco de dados local** (`db.sqlite3`).
*   Se o site já estiver rodando com dados reais de usuários, e você fizer o passo 3 acima, **você vai apagar o banco real e colocar o do seu computador no lugar**.
*   **PARE E PENSE:** Se o site tem dados importantes, **delete** o arquivo `db.sqlite3` de dentro do arquivo ZIP antes de fazer o upload para o site. Assim você atualiza só o código e preserva os dados.
