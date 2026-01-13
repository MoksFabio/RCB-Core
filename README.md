# RCB - Remuneração, Custos e Bilhetagem

Bem-vindo ao **Portal RCB**, uma plataforma desenvolvida para modernizar e redefinir a eficiência dos processos analíticos do **STPP/RMR** (Sistema de Transporte Público de Passageiros da Região Metropolitana do Recife), no âmbito do **Grande Recife Consórcio de Transporte**.

## Missão e Objetivo

O propósito central do RCB é transformar tarefas complexas, que anteriormente consumiam horas ou até dias de trabalho manual da GECO (Gerência de Contratos e Concessão), em operações automáticas, rápidas e robustas.

O sistema visa substituir processos antigos e manuais por soluções digitais que garantem:
*   **Agilidade:** Relatórios que levavam horas são gerados em menos de 30 segundos.
*   **Confiabilidade:** Eliminação de erros humanos em cálculos complexos e repetitivos.
*   **Inovação:** Criação de ferramentas inéditas, como a automação de relatórios via *ExcelJS* e *ReportLab*.

---

## Funcionalidades Detalhadas

O sistema é composto por módulos integrados, cada um responsável por uma faceta crítica da gestão de transportes. Abaixo está o detalhamento técnico e funcional de cada módulo atualmente implementado:

### 1. Cota de Óleo Diesel (cota_oleo_diesel.js)
Este é um dos módulos mais críticos do sistema, responsável por automatizar o cálculo da cota de combustível para as operadoras. A lógica, originalmente em Python (Pandas), foi totalmente portada para **JavaScript** para execução rápida no *client-side*.

*   **Processamento de Dados:**
    *   Consolidação automática de arquivos de "Quilometragem Passada" vs "Quilometragem Atual".
    *   Tratamento de dados brutos das planilhas CNO e MOB.
    *   Algoritmos inteligentes para identificação de linhas novas, excluídas ou com divergência de quilometragem.
*   **Cálculos Complexos:**
    *   Aplicação de rendimentos específicos por tipo de veículo (ex: *Padron com Ar*, *Articulado*, *Micro Urbano*).
    *   Cálculo de crédito presumido de ICMS conforme o Convênio 21/2023.
*   **Geração de Relatórios (Excel):**
    *   Gera um arquivo Excel (.xlsx) completo contendo 7 abas detalhadas: *Km Prog, Rendimento PCO, Cota do Mês, Cálculos (CNO/MOB), Rateamento e SEFAZ*.
    *   Formatação visual avançada (bordas, cores, mesclagem de células) idêntica aos relatórios oficiais.

### 2. Gestão de Ouvidorias e Demandas (ouvidorias.js)
Módulo dedicado à análise de passageiros e geração de relatórios para a Ouvidoria, reduzindo drasticamente o trabalho manual no SEI.

*   **Assistentes de Relatório:**
    *   **Demanda de Linha:** Gera relatórios analíticos para Permissionárias, Concessionárias e STPP/RMR.
    *   **Cálculo de Média:** Processa médias de passageiros e passageiros equivalentes (Média dos últimos meses).
*   **Tecnologia:**
    *   Uso intensivo da biblioteca `ExcelJS` para criar planilhas complexas diretamente no navegador.
    *   Suporte a *Drag-and-Drop* para upload de múltiplos arquivos de texto (.txt) e validação de layout.
*   **Funcionalidades de Exportação:**
    *   Geração de matrizes de resumo ("Resumo Geral por Empresa").
    *   Divisão de dados por quinzena.
    *   Inclusão automática de logotipos (CTM/RCB) nos relatórios gerados.

### 3. Análise de Congestionamento e Viagens (congestionamento.js)
Ferramenta para auditoria e validação de viagens realizadas versus programadas.

*   **Validação Cruzada:**
    *   Processa arquivos `.txt` de viagens e compara com tabelas de "Viagens Extras" e "Reduções" (formato `.ods`).
    *   Calcula o "Saldo" de viagens e o "Valor Final" a ser pago ou descontado.
*   **Lógica de Negócio:**
    *   Identifica déficits de viagens (Programado > Realizado).
    *   Aplica compensações automáticas baseadas em viagens extras validadas.
*   **Saída de Dados:**
    *   Gera arquivos `.txt` formatados para importação em sistemas legados (EP - Excesso de Passageiros/Viagens).
    *   Gera relatórios gerenciais em Excel com formatação condicional para facilitar a leitura de auditores.

### 4. Controle de Acesso e Usuários (aprovar_registros_modal.js)
Gestão administrativa segura para controle de acesso ao portal.

*   **Fluxo de Aprovação:** Novos registros entram como "Pendentes". Administradores podem Aprovar, Rejeitar ou Remover usuários.
*   **Interatividade em Tempo Real:**
    *   Ações na tabela (aprovar/remover) atualizam os KPIs do Dashboard (ex: *Usuários Ativos*, *Registros Pendentes*) instantaneamente, sem recarregar a página.
    *   Transição dinâmica de linhas entre as tabelas de "Pendentes" e "Aprovados".
*   **Segurança:** Modais de confirmação robustos para evitar ações destrutivas acidentais.

### 5. Portal de Conexões e Monitoramento (portal_de_conexoes.js)
Dashboard para monitoramento da saúde dos serviços e APIs integradas.

*   **Status de Serviços:** Visualização semafórica (Verde/Operando, Amarelo/Instável, Vermelho/Offline).
*   **Gestão Dinâmica:** Interface para adicionar, editar ou remover serviços monitorados via API backend.

---

## Tecnologias Utilizadas

*   **Backend:** Python (Django) - Robustez e segurança na gestão de dados.
*   **Frontend:**
    *   **HTML5 & CSS3:** Estrutura semântica e estilização moderna.
    *   **TailwindCSS:** Design responsivo e ágil.
    *   **JavaScript (ES6+):** Lógica complexa no lado do cliente.
    *   **ExcelJS / SheetJS:** Motores poderosos para manipulação de planilhas.
    *   **Animate.css:** Micro-interações e feedback visual fluido.

## Equipe e Créditos

Este projeto é fruto da dedicação e conhecimento técnico da equipe:

*   **Fábio Silva de Lima**
*   **Pedro Rodrigues Gonçalves**
*   **José Augusto Rocha**
*   **Luan Ferreira**

---
*Gerado automaticamente para documentação do projeto RCB.*
