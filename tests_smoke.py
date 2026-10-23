"""Teste rápido manual do armazenamento: python tests_smoke.py"""
from pathlib import Path
from tempfile import TemporaryDirectory

from src.storage import ExcelStore

with TemporaryDirectory() as td:
    path = Path(td) / "teste.xlsx"
    store = ExcelStore(path=path, template_path=None)
    store.create_ticket({
        "ticket_id": "TEST-001", "loja": "Loja de Teste", "tecnico": "Técnico Teste",
        "prioridade": "Alta", "status": "Aberto", "motivo_pendencia": "Outros",
        "ultima_atualizacao": "Registro de teste", "proxima_acao": "Validar fluxo",
        "responsavel_proxima_acao": "NOC", "turno_atual": "Manhã",
    }, "teste.operador")
    assert len(store.list_records("Chamados")) == 1
    changed = store.update_ticket("TEST-001", {"status": "Em atendimento"}, "teste.operador", "Teste de atualização")
    assert changed == 1
    passage = store.record_handover(["TEST-001"], "Manhã", "Tarde", "teste.operador", "Teste de passagem")
    assert passage.startswith("PAS-")
    assert store.list_records("Chamados")[0]["turno_atual"] == "Tarde"
    assert len(store.list_records("Historico")) >= 3
    assert len(store.list_records("Passagens")) == 1
    print("OK: criação, atualização, auditoria e passagem de turno.")
