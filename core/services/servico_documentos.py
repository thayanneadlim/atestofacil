from datetime import datetime, time
from decimal import Decimal

from django.core.exceptions import ValidationError
from django.db import transaction
from django.utils import timezone

from core.models import SequenciaAnual
from core.services.numeracao import gerar_proximo_numero
from core.services.word_template import gerar_documento_word
from core.services.templates_texto import (
    adicionar_dias_uteis,
    extrair_n_doc_sem_ano,
    extrair_nome_curto,
    formatar_cpf,
    formatar_dia_ano_doc,
    formatar_dias_trabalhados,
    formatar_horarios_trabalhados,
    formatar_observacoes,
    renderizar_texto,
)
from documentos.models import Documento, DocumentoVersao
from servicos.models import (
    DiaTrabalhado,
    Servico,
    calcular_quantidade_horas,
    montar_resumo_pagamento,
)


class ServicoDocumentoError(Exception):
    pass


def montar_contexto_de_servico(servico: Servico) -> dict:
    """Monta o contexto de variáveis a partir de um serviço já salvo."""
    documento = servico.documentos.order_by("id").first()
    data_documento = documento.data_documento if documento else timezone.localdate()
    dias = list(servico.dias_trabalhados.all())
    return _montar_contexto(
        servico.pessoa,
        servico.empresa,
        servico,
        dias,
        servico.numero_documento,
        data_documento,
    )


def _normalizar_horario(valor) -> time:
    if isinstance(valor, time):
        return valor
    if isinstance(valor, str):
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(valor, fmt).time()
            except ValueError:
                continue
    raise ServicoDocumentoError("Horário inválido. Use o formato HH:MM.")


def _montar_contexto(pessoa, empresa, servico, dias, numero_formatado, data_documento):
    total = sum((d.quantidade_horas for d in dias), Decimal("0"))
    dias_uteis = servico.dias_uteis_vencimento() if hasattr(servico, "dias_uteis_vencimento") else 1
    data_vencimento = adicionar_dias_uteis(data_documento, dias_uteis)
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
        "data_vencimento": data_vencimento.strftime("%d/%m/%Y"),
        "empresa": empresa.nome,
        "dias_trabalhados": formatar_dias_trabalhados(dias),
        "horarios_trabalhados": formatar_horarios_trabalhados(dias),
        "observacoes": formatar_observacoes(dias),
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
            "assinatura": getattr(servico, "assinatura", ""),
            "data_servico": servico.data_servico.isoformat(),
        },
        "dias": [
            {
                "data": d.data.isoformat(),
                "hora_inicio": d.hora_inicio.strftime("%H:%M"),
                "hora_fim": d.hora_fim.strftime("%H:%M"),
                "quantidade_horas": str(d.quantidade_horas),
                "observacao": d.observacao or "",
            }
            for d in dias
        ],
        "textos": {tipo: t.id for tipo, t in textos_por_tipo.items()},
    }


def _validar_entrada(dias_trabalhados, textos_por_tipo):
    if "DOCUMENTO_1" not in textos_por_tipo or textos_por_tipo.get("DOCUMENTO_1") is None:
        raise ServicoDocumentoError("O Texto Documento 1 é obrigatório.")
    for tipo, texto in textos_por_tipo.items():
        if texto is None:
            raise ServicoDocumentoError(f"Texto padrão inválido para {tipo}.")
    if not dias_trabalhados:
        raise ServicoDocumentoError("Informe ao menos um dia trabalhado.")
    datas = [item["data"] for item in dias_trabalhados]
    if len(datas) != len(set(datas)):
        raise ServicoDocumentoError("Não é permitido dia duplicado no mesmo serviço.")
    for item in dias_trabalhados:
        try:
            inicio = _normalizar_horario(item.get("hora_inicio"))
            fim = _normalizar_horario(item.get("hora_fim"))
            calcular_quantidade_horas(inicio, fim)
        except (ServicoDocumentoError, ValidationError) as exc:
            if isinstance(exc, ValidationError):
                msg = "; ".join(
                    e if isinstance(e, str) else str(e)
                    for errs in exc.message_dict.values()
                    for e in (errs if isinstance(errs, list) else [errs])
                ) if hasattr(exc, "message_dict") else str(exc)
                raise ServicoDocumentoError(msg) from exc
            raise


def _criar_dias(servico, dias_trabalhados):
    dias = []
    for item in dias_trabalhados:
        inicio = _normalizar_horario(item["hora_inicio"])
        fim = _normalizar_horario(item["hora_fim"])
        dias.append(
            DiaTrabalhado.objects.create(
                servico=servico,
                data=item["data"],
                hora_inicio=inicio,
                hora_fim=fim,
                observacao=item.get("observacao") or "",
            )
        )
    return dias


@transaction.atomic
def criar_servico_com_documentos(
    *,
    pessoa,
    empresa,
    motivo_contratacao: str,
    funcao: str,
    data_servico,
    dias_trabalhados: list[dict],
    textos_por_tipo: dict,
    assinatura: str = Servico.Assinatura.WALQUIRIA,
    usuario=None,
) -> Servico:
    _validar_entrada(dias_trabalhados, textos_por_tipo)

    numero, ano, numero_formatado = gerar_proximo_numero(SequenciaAnual.Tipo.SERVICO)
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
        assinatura=assinatura,
    )

    dias = _criar_dias(servico, dias_trabalhados)

    contexto = _montar_contexto(
        pessoa, empresa, servico, dias, numero_formatado, data_documento
    )
    snap = _snapshot(pessoa, empresa, servico, dias, textos_por_tipo)

    for tipo, texto in textos_por_tipo.items():
        conteudo = renderizar_texto(texto.conteudo, contexto)
        doc = Documento.objects.create(
            remessa=None,
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
        arquivo = gerar_documento_word(
            contexto=contexto,
            numero_documento=numero_formatado,
            versao=1,
        )
        doc.arquivo_pdf.save(arquivo.name, arquivo, save=True)
        DocumentoVersao.objects.create(
            documento=doc,
            versao=1,
            conteudo_renderizado=conteudo,
            arquivo_pdf=doc.arquivo_pdf.name if doc.arquivo_pdf else "",
            snapshot_dados=snap,
            motivo_alteracao="Criação inicial",
            criado_por=usuario,
        )

    if servico.documentos.filter(tipo_documento="DOCUMENTO_1").count() != 1:
        raise ServicoDocumentoError("Falha ao gerar o documento obrigatório do serviço.")

    return servico


@transaction.atomic
def editar_servico_com_documentos(
    *,
    servico: Servico,
    motivo_contratacao: str,
    funcao: str,
    data_servico,
    dias_trabalhados: list[dict],
    textos_por_tipo: dict,
    assinatura: str | None = None,
    usuario=None,
    motivo_alteracao: str = "Edição do serviço",
) -> Servico:
    _validar_entrada(dias_trabalhados, textos_por_tipo)

    servico.motivo_contratacao = motivo_contratacao
    servico.funcao = funcao
    servico.data_servico = data_servico or servico.data_servico
    if assinatura:
        servico.assinatura = assinatura
    servico.save()

    servico.dias_trabalhados.all().delete()
    dias = _criar_dias(servico, dias_trabalhados)

    data_documento = timezone.localdate()
    contexto = _montar_contexto(
        servico.pessoa,
        servico.empresa,
        servico,
        dias,
        servico.numero_documento,
        data_documento,
    )
    snap = _snapshot(servico.pessoa, servico.empresa, servico, dias, textos_por_tipo)

    for tipo, texto in textos_por_tipo.items():
        doc = servico.documentos.filter(tipo_documento=tipo).first()
        if doc is None:
            doc = Documento.objects.create(
                remessa=None,
                tipo_documento=tipo,
                numero_documento=servico.numero_documento,
                pessoa=servico.pessoa,
                servico=servico,
                texto_padrao=texto,
                conteudo_renderizado="",
                data_documento=data_documento,
                versao_atual=0,
                snapshot_dados={},
            )

        doc.versao_atual += 1
        doc.texto_padrao = texto
        doc.conteudo_renderizado = renderizar_texto(texto.conteudo, contexto)
        doc.data_documento = data_documento
        doc.snapshot_dados = snap
        arquivo = gerar_documento_word(
            contexto=contexto,
            numero_documento=servico.numero_documento,
            versao=doc.versao_atual,
        )
        if doc.arquivo_pdf:
            doc.arquivo_pdf.delete(save=False)
        doc.arquivo_pdf.save(arquivo.name, arquivo, save=False)
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

    return servico
