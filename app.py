from __future__ import annotations

import os
from datetime import date, datetime, time
from pathlib import Path

import pandas as pd
import streamlit as st

from src.config import MOTIVOS, PRIORIDADES, STATUS, STATUS_TERMINAIS, TURNOS
from src.storage import ExcelStore, combine_local, now_local

st.set_page_config(page_title="Passagem de Turno | NOC & Field", page_icon="🔁", layout="wide")

st.markdown("""
<style>
  .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
  div[data-testid="stMetric"] {background: rgba(100,116,139,.08); padding: 12px 14px; border-radius: 10px;}
  .app-subtitle {color: #64748b; margin-top: -0.5rem;}
</style>
""", unsafe_allow_html=True)


def build_store() -> ExcelStore:
    configured = os.getenv("NOC_FIELD_XLSX", "").strip()
    path = Path(configured) if configured else None
    return ExcelStore(path=path) if path else ExcelStore()


@st.cache_resource
def get_store() -> ExcelStore:
    return build_store()


store = get_store()

st.title("🔁 Passagem de Turno Inteligente")
st.markdown('<div class="app-subtitle">NOC & Field Service • Registro operacional, continuidade e rastreabilidade</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Operação")
    operator = st.text_input("Identificação do operador", key="operator_name", placeholder="Ex.: matrícula ou nome corporativo")
    st.caption("No MVP, esse campo identifica o registro, mas não autentica o usuário.")
    st.divider()
    st.markdown("**Base de dados**")
    st.caption(str(store.path))
    st.caption("Última consulta: " + now_local().strftime("%d/%m/%Y %H:%M"))

try:
    tickets = store.list_records("Chamados")
    history = store.list_records("Historico")
    handovers = store.list_records("Passagens")
except Exception as exc:
    st.error(f"Não foi possível consultar a base Excel: {exc}")
    st.stop()

now = now_local()
open_tickets = [t for t in tickets if t.get("status") not in STATUS_TERMINAIS]
critical = [t for t in open_tickets if t.get("prioridade") in {"Crítica", "Alta"}]
overdue = []
for t in open_tickets:
    due = t.get("prazo_sla")
    if isinstance(due, datetime) and due < now:
        overdue.append(t)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Pendências abertas", len(open_tickets))
k2.metric("Prioridade alta/crítica", len(critical))
k3.metric("SLA vencido", len(overdue))
k4.metric("Passagens registradas", len(handovers))


def require_operator() -> bool:
    if not operator.strip():
        st.error("Preencha a identificação do operador na barra lateral antes de registrar alterações.")
        return False
    return True


def value_or_blank(value):
    return value if value is not None else ""


def as_date(value, fallback=None):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return fallback or date.today()


def as_time(value, fallback=None):
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    return fallback or time(18, 0)


def ticket_label(ticket_id: str) -> str:
    rec = next((x for x in tickets if str(x.get("ticket_id")) == str(ticket_id)), {})
    return f"{ticket_id} | {rec.get('loja') or 'Loja não informada'} | {rec.get('status') or 'Sem status'}"

# As quatro frentes operacionais ficam juntas nesta primeira versão.
tab_new, tab_list, tab_handover, tab_history = st.tabs([
    "➕ Novo chamado", "📋 Consultar e atualizar", "🔄 Passagem de turno", "🕘 Histórico de auditoria"
])

with tab_new:
    st.subheader("Registrar pendência de Field Service")
    st.caption("Preencha o essencial. Os campos de data/hora de criação e auditoria são registrados automaticamente.")
    with st.form("new_ticket_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            ticket_id = st.text_input("ID do chamado *", placeholder="Ex.: INC123456")
            loja = st.text_input("Loja / unidade *", placeholder="Ex.: Loja 025")
            tecnico = st.text_input("Técnico responsável", placeholder="Nome do técnico")
        with c2:
            prioridade = st.selectbox("Prioridade *", PRIORIDADES, index=2)
            status = st.selectbox("Status inicial", [s for s in STATUS if s not in STATUS_TERMINAIS], index=0)
            turno = st.selectbox("Turno atual", TURNOS, index=0)
        with c3:
            motivo = st.selectbox("Motivo / impedimento *", MOTIVOS)
            define_sla = st.checkbox("Informar prazo de SLA", value=True)
            sla_date = st.date_input("Data limite do SLA", value=date.today(), disabled=not define_sla)
            sla_time = st.time_input("Hora limite do SLA", value=time(18, 0), disabled=not define_sla, step=300)
        ultima = st.text_area("Última atualização *", placeholder="O que já foi feito / última informação confirmada?")
        c4, c5 = st.columns(2)
        with c4:
            proxima = st.text_area("Próxima ação *", placeholder="Qual é a ação objetiva para o próximo turno?")
        with c5:
            responsavel = st.text_input("Responsável pela próxima ação", placeholder="Ex.: NOC - turno tarde")
        observacao = st.text_input("Observação complementar (opcional)")
        submitted = st.form_submit_button("Salvar chamado", type="primary", use_container_width=True)
    if submitted:
        if not require_operator():
            pass
        elif not ticket_id.strip() or not loja.strip() or not ultima.strip() or not proxima.strip():
            st.error("Preencha os campos obrigatórios: ID, loja, última atualização e próxima ação.")
        else:
            try:
                store.create_ticket({
                    "ticket_id": ticket_id.strip(), "loja": loja.strip(), "tecnico": tecnico.strip(),
                    "prioridade": prioridade, "status": status, "motivo_pendencia": motivo,
                    "data_abertura": now_local(),
                    "prazo_sla": combine_local(sla_date, sla_time) if define_sla else None,
                    "ultima_atualizacao": ultima.strip(), "proxima_acao": proxima.strip(),
                    "responsavel_proxima_acao": responsavel.strip(), "turno_atual": turno,
                    "resolucao": observacao.strip() if status in STATUS_TERMINAIS else "",
                }, operator)
                st.success(f"Chamado {ticket_id.strip()} registrado com sucesso.")
                st.rerun()
            except Exception as exc:
                st.error(str(exc))

with tab_list:
    st.subheader("Consulta e atualização de pendências")
    if not tickets:
        st.info("Ainda não há chamados cadastrados. Use a aba 'Novo chamado' para registrar a primeira pendência.")
    else:
        f1, f2, f3 = st.columns([1.4, 1, 1])
        with f1:
            query = st.text_input("Buscar por ID, loja ou técnico", key="ticket_query")
        with f2:
            statuses = st.multiselect("Status", STATUS, default=[s for s in STATUS if s not in STATUS_TERMINAIS], key="ticket_status_filter")
        with f3:
            priorities = st.multiselect("Prioridade", PRIORIDADES, default=PRIORIDADES, key="ticket_priority_filter")
        filtered = []
        q = query.casefold().strip()
        for rec in tickets:
            if statuses and rec.get("status") not in statuses:
                continue
            if not statuses:
                continue
            if priorities and rec.get("prioridade") not in priorities:
                continue
            if not priorities:
                continue
            searchable = " ".join(str(rec.get(k) or "") for k in ["ticket_id", "loja", "tecnico"]).casefold()
            if q and q not in searchable:
                continue
            filtered.append(rec)
        show_cols = ["ticket_id", "loja", "tecnico", "prioridade", "status", "motivo_pendencia", "prazo_sla", "turno_atual", "responsavel_proxima_acao"]
        table_data = [{k: rec.get(k) for k in show_cols} for rec in filtered]
        st.dataframe(pd.DataFrame(table_data, columns=show_cols), use_container_width=True, hide_index=True)
        st.caption(f"Exibindo {len(filtered)} de {len(tickets)} chamados. Para ver chamados encerrados, inclua 'Resolvido' ou 'Cancelado' no filtro de status.")

        st.divider()
        st.markdown("#### Atualizar um chamado existente")
        ids = [str(x.get("ticket_id")) for x in tickets if x.get("ticket_id")]
        chosen = st.selectbox("Chamado para editar", ids, format_func=ticket_label, key="edit_ticket_choice")
        rec = next((x for x in tickets if str(x.get("ticket_id")) == chosen), {})
        with st.form("edit_ticket_form"):
            e1, e2, e3 = st.columns(3)
            with e1:
                e_loja = st.text_input("Loja / unidade", value=str(rec.get("loja") or ""))
                e_tecnico = st.text_input("Técnico responsável", value=str(rec.get("tecnico") or ""))
                e_status = st.selectbox("Status", STATUS, index=STATUS.index(rec.get("status")) if rec.get("status") in STATUS else 0)
            with e2:
                e_prioridade = st.selectbox("Prioridade", PRIORIDADES, index=PRIORIDADES.index(rec.get("prioridade")) if rec.get("prioridade") in PRIORIDADES else 2)
                e_motivo = st.selectbox("Motivo / impedimento", MOTIVOS, index=MOTIVOS.index(rec.get("motivo_pendencia")) if rec.get("motivo_pendencia") in MOTIVOS else len(MOTIVOS)-1)
                e_turno = st.selectbox("Turno responsável atual", TURNOS, index=TURNOS.index(rec.get("turno_atual")) if rec.get("turno_atual") in TURNOS else 0)
            with e3:
                existing_sla = rec.get("prazo_sla")
                e_define_sla = st.checkbox("Manter / definir prazo de SLA", value=isinstance(existing_sla, datetime))
                e_sla_date = st.date_input("Data limite do SLA", value=as_date(existing_sla), disabled=not e_define_sla, key="edit_sla_date")
                e_sla_time = st.time_input("Hora limite do SLA", value=as_time(existing_sla), disabled=not e_define_sla, step=300, key="edit_sla_time")
            e_ultima = st.text_area("Última atualização", value=str(rec.get("ultima_atualizacao") or ""))
            x1, x2 = st.columns(2)
            with x1:
                e_proxima = st.text_area("Próxima ação", value=str(rec.get("proxima_acao") or ""))
            with x2:
                e_responsavel = st.text_input("Responsável pela próxima ação", value=str(rec.get("responsavel_proxima_acao") or ""))
            e_resolucao = st.text_area("Resolução / observação de encerramento", value=str(rec.get("resolucao") or ""))
            e_comentario = st.text_input("Comentário da alteração (opcional)")
            save_edit = st.form_submit_button("Salvar atualização", type="primary", use_container_width=True)
        if save_edit:
            if not require_operator():
                pass
            elif not e_loja.strip() or not e_ultima.strip() or not e_proxima.strip():
                st.error("Loja, última atualização e próxima ação são obrigatórias.")
            elif e_status == "Resolvido" and not e_resolucao.strip():
                st.error("Informe a resolução antes de marcar o chamado como Resolvido.")
            else:
                try:
                    count = store.update_ticket(chosen, {
                        "loja": e_loja.strip(), "tecnico": e_tecnico.strip(), "status": e_status,
                        "prioridade": e_prioridade, "motivo_pendencia": e_motivo,
                        "turno_atual": e_turno,
                        "prazo_sla": combine_local(e_sla_date, e_sla_time) if e_define_sla else None,
                        "ultima_atualizacao": e_ultima.strip(), "proxima_acao": e_proxima.strip(),
                        "responsavel_proxima_acao": e_responsavel.strip(), "resolucao": e_resolucao.strip(),
                    }, operator, e_comentario.strip())
                    if count:
                        st.success(f"Chamado {chosen} atualizado. {count} campo(s) alterado(s); histórico registrado.")
                    else:
                        st.info("Nenhuma alteração detectada.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

with tab_handover:
    st.subheader("Registrar transferência entre turnos")
    active = [x for x in tickets if x.get("status") not in STATUS_TERMINAIS and x.get("ticket_id")]
    if not active:
        st.info("Não há pendências abertas para transferir.")
    else:
        id_list = [str(x.get("ticket_id")) for x in active]
        with st.form("handover_form"):
            selected_ids = st.multiselect("Pendências transferidas *", id_list, format_func=ticket_label, help="Selecione todos os chamados que estão sendo passados ao turno seguinte.")
            h1, h2, h3 = st.columns(3)
            with h1:
                source = st.selectbox("Turno de origem", TURNOS, index=0)
            with h2:
                target = st.selectbox("Turno de destino", TURNOS, index=1)
            with h3:
                receiver = st.text_input("Recebido por (opcional)", placeholder="Preencha apenas após confirmar o recebimento")
            summary = st.text_area("Resumo geral da passagem *", placeholder="Riscos, prioridades e orientações importantes para o turno entrante")
            confirm_handover = st.form_submit_button("Registrar passagem de turno", type="primary", use_container_width=True)
        if confirm_handover:
            if not require_operator():
                pass
            else:
                try:
                    passage_id = store.record_handover(selected_ids, source, target, operator, summary, receiver)
                    st.success(f"Passagem {passage_id} registrada para {len(selected_ids)} chamado(s). O turno responsável foi atualizado para '{target}'.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
    if handovers:
        st.divider()
        st.markdown("#### Passagens recentes")
        recent = sorted(handovers, key=lambda x: x.get("data_hora") or datetime.min, reverse=True)[:20]
        st.dataframe(pd.DataFrame(recent), use_container_width=True, hide_index=True)

with tab_history:
    st.subheader("Trilha de auditoria")
    st.caption("O histórico registra criação, atualização de campos e transferência de responsabilidade.")
    hx1, hx2 = st.columns([2, 1])
    with hx1:
        h_query = st.text_input("Filtrar por ID do chamado", key="history_ticket_filter")
    with hx2:
        event_options = sorted({str(x.get("tipo_evento")) for x in history if x.get("tipo_evento")})
        event_filter = st.selectbox("Tipo de evento", ["Todos"] + event_options)
    shown = history
    if h_query.strip():
        shown = [x for x in shown if h_query.strip().casefold() in str(x.get("ticket_id") or "").casefold()]
    if event_filter != "Todos":
        shown = [x for x in shown if x.get("tipo_evento") == event_filter]
    shown = sorted(shown, key=lambda x: x.get("data_hora") or datetime.min, reverse=True)[:500]
    if shown:
        columns = ["data_hora", "ticket_id", "usuario", "tipo_evento", "campo", "valor_anterior", "valor_novo", "comentario"]
        st.dataframe(pd.DataFrame([{c: row.get(c) for c in columns} for row in shown]), use_container_width=True, hide_index=True)
        st.caption(f"Exibindo até 500 eventos mais recentes. Eventos armazenados na base: {len(history)}.")
    else:
        st.info("Não há eventos que correspondam aos filtros selecionados.")

st.divider()
st.caption("MVP operacional • Excel como armazenamento inicial • Faça backup do arquivo em data/ e evite editá-lo manualmente enquanto o aplicativo estiver em execução.")
