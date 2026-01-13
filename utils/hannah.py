import datetime
import pandas as pd
import re
from django.db.models import Q
# Import Django models. 
# Note: This file is imported by views.py, so Django environment is ready.
try:
    from portal.models import User, SystemService, Manifestacao, Compromisso
except ImportError:
    # Fallback for testing outside Django context if needed, though mostly this runs in Django
    pass

def get_hannah_response(user_message, username="Usuário"):
    """
    Processa a mensagem do usuário e retorna um dicionário com:
    - 'text': A resposta em texto da Hannah.
    - 'action': (Opcional) Uma string de comando para o frontend executar.
    """
    user_message = user_message.lower().strip()
    response = {"text": "", "action": None}
    
    # --- 1. SAUDAÇÕES ---
    if 'olá' in user_message or 'oi ' in user_message or user_message == 'oi':
        response["text"] = f"Olá, {username}! Sou a Hannah, sua assistente inteligente. Posso verificar status do sistema, buscar linhas, fazer cálculos ou abrir ferramentas para você. Como ajudo?"
        return response

    elif 'ajuda' in user_message or 'o que você faz' in user_message:
        response["text"] = (
            "Aqui está o que eu posso fazer por você:\n\n"
            "🔍 **Consultas:**\n"
            "• 'Código da linha [Nome]'\n"
            "• 'Status do sistema' ou 'Pendências'\n"
            "• 'Calcular 100 / 4' ou 'Converter 90 min em horas'\n\n"
            "🚀 **Ações (Posso abrir para você):**\n"
            "• 'Abrir congestionamento'\n"
            "• 'Minha agenda'\n"
            "• 'Novo registro'\n\n"
            "E claro, dizer as horas e informações sobre as ferramentas!"
        )
        return response

    # --- 2. COMANDOS DE AÇÃO (NAVEGAÇÃO) ---
    actions_map = {
        'congestionamento': 'OPEN_MODAL_CONGESTIONAMENTO',
        'passageiro': 'NAVIGATE_PASSAGEIRO',
        'integrado': 'NAVIGATE_PASSAGEIRO',
        'cota': 'OPEN_MODAL_COTA',
        'oleo': 'OPEN_MODAL_COTA',
        'diesel': 'OPEN_MODAL_COTA',
        'demanda': 'OPEN_MODAL_OUVIDORIAS',
        'ouvidoria': 'OPEN_MODAL_OUVIDORIAS',
        'sabe': 'NAVIGATE_SABE',
        'agenda': 'SCROLL_TO_AGENDA',
        'evento': 'SCROLL_TO_AGENDA',
        'bloco': 'SCROLL_TO_NOTAS',
        'notas': 'SCROLL_TO_NOTAS',
        'aprovar': 'OPEN_MODAL_APROVAR',
        'registros': 'OPEN_MODAL_APROVAR',
        'pendentes': 'OPEN_MODAL_APROVAR',
        'frota': 'OPEN_MODAL_FROTA',
        'remuneracao': 'OPEN_MODAL_PARAMETROS',
    }
    
    if 'abrir' in user_message or 'ir para' in user_message or 'mostrar' in user_message:
        for key, action_code in actions_map.items():
            if key in user_message:
                response["text"] = f"Abrindo {key} para você..."
                response["action"] = action_code
                return response

    # --- 3. STATUS DO SISTEMA (Django Models) ---
    if 'status' in user_message or 'sistema' in user_message and 'como' in user_message:
        try:
            services = SystemService.objects.all()
            offline = services.filter(status='offline').count()
            instavel = services.filter(status='instavel').count()
            
            if offline == 0 and instavel == 0:
                response["text"] = "✅ Todos os sistemas estão operando normalmente!"
            else:
                details = []
                if offline > 0: details.append(f"{offline} offline")
                if instavel > 0: details.append(f"{instavel} instável")
                response["text"] = f"⚠️ Atenção: Há serviços com problemas ({', '.join(details)}). Verifique o widget de status."
            return response
        except Exception as e:
            response["text"] = "Não consegui verificar o status dos serviços no momento."
            return response

    if 'pendências' in user_message or 'pendentes' in user_message or 'aprovar' in user_message:
        try:
            pending_count = User.objects.filter(status='pendente').count()
            if pending_count == 0:
                response["text"] = "Não há novos registros de usuários pendentes de aprovação."
            else:
                response["text"] = f"Há **{pending_count}** usuário(s) aguardando aprovação."
                response["action"] = "NAVIGATE_APPROVE" # Sugere ir lá
            return response
        except:
            pass

    # --- 4. DATA E HORA ---
    if 'que horas são' in user_message or 'hora atual' in user_message:
        now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-3)))
        response["text"] = f"Agora são {now.strftime('%H:%M')} em Recife."
        return response

    # --- 5. CÁLCULO E CONVERSÃO SIMPLES ---
    # Convert minutes to hours decimal: "converter 150 min"
    conv_match = re.search(r'converter (\d+)[\s]*min', user_message)
    if conv_match:
        minutes = int(conv_match.group(1))
        hours = minutes / 60
        response["text"] = f"{minutes} minutos equivalem a **{hours:.2f} horas**."
        return response
    
    # Basic Math: "calcular 10 + 20"
    if 'calcular' in user_message or 'quanto é' in user_message:
        try:
            # Extrai apenas números e operadores básicos para segurança
            expression = re.sub(r'[^0-9\+\-\*\/.]', '', user_message.split(' ', 1)[1])
            if expression:
                result = eval(expression)
                response["text"] = f"O resultado é **{result}**."
                return response
        except:
            response["text"] = "Não entendi a conta. Tente algo como 'calcular 100 / 4'."
            return response

    # --- 6. INFO FERRAMENTAS (Mantendo o original, mas simplificado) ---
    tool_infos = {
        'congestionamento': "O 'Congestionamento' compara viagens programadas x realizadas. Gera arquivos .txt e relatórios de saldo.",
        'passageiro': "O 'Passageiro Integrado' gera relatórios PDF consolidados a partir de arquivos de bilhetagem.",
        'cota': "A 'Cota de Óleo' audita a quilometragem e combustível entre dois meses de referência.",
        'demanda média': "Calcula a média de passageiros (DUT, SAB, DOM) para um período.",
        'demanda': "Gera relatórios detalhados de demanda por linha para ouvidoria.", # Demanda genérica se não for média
        'sabe': "O 'SABE' compara dados de remuneração (.txt) com catraca (.dbf) para auditoria."
    }
    
    for key, info in tool_infos.items():
        if key in user_message:
            response["text"] = info
            return response

    # --- 7. BUSCA DE LINHA (Código ou Nome) ---
    # Verifica se é busca por código (somente números)
    try:
        from django.conf import settings
        file_path = settings.BASE_DIR.parent / 'frontend' / 'src' / 'static' / 'linhas.xlsx'
        coluna_codigo = 'CÓDIGO LINHA'
        coluna_nome_linha = 'NOME LINHA'
        coluna_operadora = 'OPERADOR'

        # Carrega DF (poderia ser carregado globalmente para performance, mas aqui garante atualização)
        # Otimização: ler apenas se parecer uma busca de linha
        
        # Padrão busca nome: "código da linha x" ou "linha x" com texto
        is_name_search = 'código' in user_message or 'linha' in user_message
        
        # Padrão código direto
        is_code_search = user_message.isdigit()

        if is_code_search or is_name_search:
            try:
                df_linhas = pd.read_excel(file_path, engine='openpyxl', dtype={coluna_codigo: str})
                df_linhas.columns = df_linhas.columns.str.strip()
            except FileNotFoundError:
                response["text"] = "Erro: Arquivo de linhas não encontrado no servidor."
                return response

            found_row = None
            
            if is_code_search:
                # Busca exata pelo código
                matches = df_linhas[df_linhas[coluna_codigo] == user_message]
                if not matches.empty:
                    found_row = matches
                else:
                    response["text"] = f"Não encontrei nenhuma linha com o código {user_message}."
                    return response
            
            elif is_name_search:
                # Extrai o termo de busca. Ex: "código da linha barra de jangada" -> "barra de jangada"
                search_term = user_message.replace('código', '').replace('da linha', '').replace('linha', '').strip()
                
                if len(search_term) < 3:
                     # Evita buscas muito curtas se não for numérico
                     if not user_message.isdigit():
                         pass 
                else:
                    # Busca textual case-insensitive
                    matches = df_linhas[df_linhas[coluna_nome_linha].astype(str).str.contains(search_term, case=False, na=False)]
                    
                    if matches.empty:
                         response["text"] = f"Não encontrei linhas com o nome '{search_term}'."
                         return response
                    elif len(matches) > 1:
                        # Retorna lista se houver poucos, ou pede para refinar
                        if len(matches) <= 5:
                            msg = f"Encontrei {len(matches)} linhas:\n"
                            for _, row in matches.iterrows():
                                msg += f"• **{row[coluna_codigo]}**: {row[coluna_nome_linha]}\n"
                            response["text"] = msg
                            return response
                        else:
                            response["text"] = f"Encontrei {len(matches)} linhas com '{search_term}'. Seja mais específico."
                            return response
                    else:
                        found_row = matches

            # Formata resposta se encontrou uma linha única
            if found_row is not None:
                nome = found_row.iloc[0][coluna_nome_linha]
                cod = found_row.iloc[0][coluna_codigo]
                operadoras = found_row[coluna_operadora].unique().tolist()
                
                resposta_final = f"🚍 **Linha {cod}**: {nome}\n"
                if len(operadoras) > 1:
                    resposta_final += "Operadoras: " + ", ".join(operadoras)
                else:
                    resposta_final += f"Operadora: {operadoras[0]}"
                
                response["text"] = resposta_final
                return response
                
    except Exception as e:
        print(f"Hannah Error: {e}") # Log no console do servidor
        # Não retorna erro explícito pro usuário se não for certeza que era uma busca de linha,
        # deixa cair no "não entendi" final.

    # --- FINAL: NÃO ENTENDI ---
    response["text"] = "Desculpe, não entendi. Tente 'ajuda' para ver o que posso fazer."
    return response