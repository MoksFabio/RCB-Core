
# ☁️ Manual de Sincronização (Git)

Checklist padrão para manter o projeto atualizado e salvo na nuvem.

## 1. Configuração de Novo Ambiente
*Executar apenas na primeira vez ao configurar o projeto em uma nova máquina.*

1.   **Instale o Git**: [Baixar Git](https://git-scm.com/download/win)
2.  Abra o terminal na pasta desejada.
3.  **Clone o projeto**:
    ```bash
    git clone https://github.com/rcb-remuneracao-custos-e-bilhetagem/RCB-DJANGO.git
    ```

---

## 2. Fluxo de Desenvolvimento

### 🟢 Ao Iniciar (Sincronizar)
Sempre baixe a versão mais recente do servidor antes de começar a codificar.
```bash
git pull
```

### 🔴 Ao Finalizar (Salvar e Enviar)
Envie suas alterações para o repositório remoto para salvar seu progresso.
```bash
git add .
git commit -m "Rotina de atualização"
git push
```

---

## 3. Recuperação e Histórico (Emergência)

### 🧹 Descartar Alterações Não Salvas
Se você alterou arquivos mas **ainda não fez o commit** e quer cancelar tudo (voltar ao estado limpo):
```bash
git checkout .
```

### 🕰️ Voltar para uma Versão Antiga
Se você precisa ver ou restaurar como o código estava no passado:

1.  **Liste o histórico** para achar o código da versão (Hash):
    ```bash
    git log --oneline
    ```
2.  **Volte no tempo** (modo somente leitura):
    ```bash
    git checkout <codigo_do_hash>
    ```
3.  **Retorne ao presente** (para continuar trabalhando):
    ```bash
    git checkout main
    ```

---

## 4. Solução de Problemas (Login/Senha)

### 🔐 Erro "Repository not found" ou Permissão Negada
Se você tiver certeza que o repositório existe, mas o Git insistir que não (geralmente por conflito de login salvo no Windows):

Use este comando para **forçar** o login manual:
```bash
git -c credential.helper= push
```

### 📥 Erro ao Baixar (Clone)
Se acontecer o mesmo erro ("Repository not found") ao tentar baixar o projeto pela primeira vez:
```bash
git -c credential.helper= clone https://github.com/rcb-remuneracao-custos-e-bilhetagem/RCB-DJANGO.git
```
