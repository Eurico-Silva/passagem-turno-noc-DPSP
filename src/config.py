from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = BASE_DIR / "templates" / "modelo_passagem_turno.xlsx"
DEFAULT_DATA_PATH = BASE_DIR / "data" / "passagem_turno.xlsx"

SHEETS = {
    "Chamados": [
        "ticket_id", "loja", "tecnico", "prioridade", "status", "motivo_pendencia",
        "data_abertura", "prazo_sla", "ultima_atualizacao", "proxima_acao",
        "responsavel_proxima_acao", "turno_atual", "criado_em", "criado_por",
        "atualizado_em", "atualizado_por", "encerrado_em", "resolucao",
    ],
    "Historico": [
        "evento_id", "ticket_id", "data_hora", "usuario", "tipo_evento",
        "campo", "valor_anterior", "valor_novo", "comentario",
    ],
    "Passagens": [
        "passagem_id", "data_hora", "turno_origem", "turno_destino", "operador",
        "quantidade_chamados", "chamados_ids", "resumo", "recebido_por", "recebido_em",
    ],
}

PRIORIDADES = ["Crítica", "Alta", "Média", "Baixa"]
STATUS = [
    "Aberto", "Em atendimento", "Aguardando loja", "Aguardando peça",
    "Aguardando técnico", "Agendado", "Resolvido", "Cancelado",
]
STATUS_TERMINAIS = {"Resolvido", "Cancelado"}
MOTIVOS = [
    "Aguardando autorização de acesso", "Aguardando loja", "Aguardando técnico",
    "Aguardando peça", "Aguardando agendamento", "Dependência de terceiro",
    "Falha técnica", "Outros",
]
TURNOS = ["Manhã", "Tarde", "Noite", "Madrugada", "Administrativo", "Outro"]
