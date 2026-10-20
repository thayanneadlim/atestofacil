from decimal import Decimal

from django.db import transaction
from django.utils import timezone

from core.models import SequenciaAnual
from core.services.numeracao import gerar_proximo_numero
from core.services.pdf import gerar_pdf_documento
from core.services.templates_texto import (
    extrair_n_doc_sem_ano,
    extrair_nome_curto,
    formatar_cpf,
    formatar_dia_ano_doc,
    formatar_dias_trabalhados,
    renderizar_texto,
)
from documentos.models import Documento, DocumentoVersao
from remessas.models import Remessa
from servicos.models import DiaTrabalhado, Servico, montar_resumo_pagamento


class RemessaError(Exception):
    pass


def _montar_contexto(pessoa, empresa, servico, dias, numero_formatado, data_documento):
    total = sum((d.quantidade_horas for d in dias), Decimal("0"))
    data_documento_fmt = data_documento.strftime("%d/%m/%Y")
    resumo_pagamento = montar_resumo_pagamento(dias)
    return {
        "nome": pessoa.nome,
        "nome_curto": extrair_nome_curto(pessoa.nome),
        "cpf": formatar_cpf(pessoa.cpf),
        "chave_pix": pessoa.chave_pix,
        "numero_documento": numero_formatado,
        "n_doc_sem_ano": extrair_n_doc_sem_ano(numero_formatado),
        "numero_rpa": getattr(servico, "numero_rpa_documento", "") or "",
        "n_rpa_sem_ano": extrair_n_doc_sem_ano(
            getattr(servico, "numero_rpa_documento", "") or ""
        ),
        "data_documento": data_documento_fmt,
        "dia_ano_doc": formatar_dia_ano_doc(data_documento),
        "empresa": empresa.nome,
        "dias_trabalhados": formatar_dias_trabalhados(dias),
        "quantidade_dias": resumo_pagamento.total_dias,
        "motivo_contratacao": servico.motivo_contratacao,
        "funcao": servico.funcao,
        "total_horas": f"{total.normalize():f}".rstrip("0").rstrip(".") if total else "0",
        "valor_total": resumo_pagamento.valor_total,
        "valor_total_formatado": resumo_pagamento.valor_total_formatado,
    }


def _snapshot(pessoa, empresa, servico, dias, textos_por_tipo):
    return {
        "pessoa": {"id": pessoa.id, "nome": pessoa.nome, "cpf": pessoa.cpf},
        "empresa": {"id": empresa.id, "nome": empresa.nome},
        "servico": {
            "motivo_contratacao": servico.motivo_contratacao,
            "funcao": servico.funcao,
            "data_servico": servico.data_servico.isoformat(),
        },
        "dias": [
            {
                "data": d.data.isoformat(),
                "quantidade_horas": str(d.quantidade_horas),
                "observacao": d.observacao or "",
            }
            for d in dias
        ],
        "textos": {tipo: t.id for tipo, t in textos_por_tipo.items()},
    }


@transaction.atomic
def criar_remessa_completa(
    *,
    pessoa,
    empresa,
    motivo_contratacao: str,
    funcao: str,
    data_servico,
    dias_trabalhados: list[dict],
    textos_por_tipo: dict,
    usuario=None,
) -> Remessa:
    if len(textos_por_tipo) != 3:
        raise RemessaError("É obrigatório selecionar exatamente 3 textos padrões.")
    for tipo in ("DOCUMENTO_1", "DOCUMENTO_2", "DOCUMENTO_3"):
        if tipo not in textos_por_tipo:
            raise RemessaError(f"Texto padrão ausente para {tipo}.")
    if not dias_trabalhados:
        raise RemessaError("Informe ao menos um dia trabalhado.")

    datas = [item["data"] for item in dias_trabalhados]
    if len(datas) != len(set(datas)):
        raise RemessaError("Não é permitido dia duplicado no mesmo serviço.")

    for item in dias_trabalhados:
        if Decimal(str(item["quantidade_horas"])) <= 0:
            raise RemessaError("Quantidade de horas deve ser maior que zero.")

    numero, ano, numero_formatado = gerar_proximo_numero(SequenciaAnual.Tipo.REMESSA)
    numero_rpa, ano_rpa, numero_rpa_formatado = gerar_proximo_numero(SequenciaAnual.Tipo.RPA)
    data_documento = timezone.localdate()

    servico = Servico.objects.create(
        numero=numero,
        ano=ano,
        numero_documento=numero_formatado,
        numero_rpa=numero_rpa,
        ano_rpa=ano_rpa,
        numero_rpa_documento=numero_rpa_formatado,
        empresa=empresa,
        pessoa=pessoa,
        data_servico=data_servico or data_documento,
        motivo_contratacao=motivo_contratacao,
        funcao=funcao,
    )

    dias = []
    for item in dias_trabalhados:
        dias.append(
            DiaTrabalhado.objects.create(
                servico=servico,
                data=item["data"],
                quantidade_horas=item["quantidade_horas"],
                observacao=item.get("observacao") or "",
            )
        )

    remessa = Remessa.objects.create(
        numero=numero,
        ano=ano,
        numero_formatado=numero_formatado,
        pessoa=pessoa,
        servico=servico,
        data_criacao=data_documento,
        status=Remessa.Status.GERADA,
    )
    servico.remessa = remessa
    servico.save(update_fields=["remessa", "updated_at"])

    contexto = _montar_contexto(
        pessoa, empresa, servico, dias, numero_formatado, data_documento
    )
    snap = _snapshot(pessoa, empresa, servico, dias, textos_por_tipo)

    for tipo, texto in textos_por_tipo.items():
        conteudo = renderizar_texto(texto.conteudo, contexto)
        doc = Documento.objects.create(
            remessa=remessa,
            tipo_documento=tipo,
            numero_documento=numero_formatado,
            pessoa=pessoa,
            servico=servico,
            texto_padrao=texto,
            conteudo_renderizado=conteudo,
            data_documento=data_documento,
            versao_atual=1,
            snapshot_dados=snap,
        )
        pdf = gerar_pdf_documento(documento=doc, titulo=texto.nome)
        doc.arquivo_pdf.save(pdf.name, pdf, save=True)
        DocumentoVersao.objects.create(
            documento=doc,
            versao=1,
            conteudo_renderizado=conteudo,
            arquivo_pdf=doc.arquivo_pdf.name if doc.arquivo_pdf else "",
            snapshot_dados=snap,
            motivo_alteracao="Criação inicial",
            criado_por=usuario,
        )

    if remessa.documentos.count() != 3:
        raise RemessaError("Falha ao gerar os 3 documentos da remessa.")

    return remessa


@transaction.atomic
def editar_remessa(
    *,
    remessa: Remessa,
    motivo_contratacao: str,
    funcao: str,
    data_servico,
    dias_trabalhados: list[dict],
    textos_por_tipo: dict,
    usuario=None,
    motivo_alteracao: str = "Edição da remessa",
) -> Remessa:
    if remessa.status == Remessa.Status.CANCELADA:
        raise RemessaError("Remessa cancelada não pode ser editada.")
    if len(textos_por_tipo) != 3:
        raise RemessaError("É obrigatório selecionar exatamente 3 textos padrões.")
    if not dias_trabalhados:
        raise RemessaError("Informe ao menos um dia trabalhado.")

    datas = [item["data"] for item in dias_trabalhados]
    if len(datas) != len(set(datas)):
        raise RemessaError("Não é permitido dia duplicado no mesmo serviço.")

    servico = remessa.servico
    servico.motivo_contratacao = motivo_contratacao
    servico.funcao = funcao
    servico.data_servico = data_servico or servico.data_servico
    servico.save()

    servico.dias_trabalhados.all().delete()
    dias = []
    for item in dias_trabalhados:
        if Decimal(str(item["quantidade_horas"])) <= 0:
            raise RemessaError("Quantidade de horas deve ser maior que zero.")
        dias.append(
            DiaTrabalhado.objects.create(
                servico=servico,
                data=item["data"],
                quantidade_horas=item["quantidade_horas"],
                observacao=item.get("observacao") or "",
            )
        )

    data_documento = timezone.localdate()
    contexto = _montar_contexto(
        remessa.pessoa,
        servico.empresa,
        servico,
        dias,
        remessa.numero_formatado,
        data_documento,
    )
    snap = _snapshot(remessa.pessoa, servico.empresa, servico, dias, textos_por_tipo)

    for tipo, texto in textos_por_tipo.items():
        doc = remessa.documentos.get(tipo_documento=tipo)
        doc.versao_atual += 1
        doc.texto_padrao = texto
        doc.conteudo_renderizado = renderizar_texto(texto.conteudo, contexto)
        doc.data_documento = data_documento
        doc.snapshot_dados = snap
        pdf = gerar_pdf_documento(documento=doc, titulo=texto.nome)
        if doc.arquivo_pdf:
            doc.arquivo_pdf.delete(save=False)
        doc.arquivo_pdf.save(pdf.name, pdf, save=False)
        doc.save()
        DocumentoVersao.objects.create(
            documento=doc,
            versao=doc.versao_atual,
            conteudo_renderizado=doc.conteudo_renderizado,
            arquivo_pdf=doc.arquivo_pdf.name if doc.arquivo_pdf else "",
            snapshot_dados=snap,
            motivo_alteracao=motivo_alteracao,
            criado_por=usuario,
        )

    remessa.status = Remessa.Status.EDITADA
    remessa.save(update_fields=["status", "updated_at"])
    return remessa
