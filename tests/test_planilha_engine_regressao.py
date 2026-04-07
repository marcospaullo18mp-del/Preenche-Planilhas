import unittest

from planilha_engine import (
    extract_indicador_geral_completo,
    extract_meta_especifica_sections,
)


class TestRegressaoVariacoesPDF(unittest.TestCase):
    def test_indicador_formato_antigo_em_linhas_separadas(self):
        lines = [
            "Alguma linha",
            "Indicador Geral de Resultado",
            "Taxa de redução de crimes por 100 mil.",
            "META ESPECÍFICA 1",
        ]
        self.assertEqual(
            extract_indicador_geral_completo(lines),
            "Taxa de redução de crimes por 100 mil.",
        )

    def test_indicador_inline_apos_titulo(self):
        lines = [
            "Indicador Geral de Resultado: Taxa Y",
            "META ESPECÍFICA 1",
        ]
        self.assertEqual(extract_indicador_geral_completo(lines), "Taxa Y")

    def test_indicador_inline_dentro_de_meta_geral(self):
        lines = [
            "Meta Geral: Reduzir índice X Indicador Geral de Resultado: Taxa Z",
            "META ESPECÍFICA 1",
        ]
        self.assertEqual(extract_indicador_geral_completo(lines), "Taxa Z")

    def test_periodicidade_com_fonte_ano_inline_padrao_antigo(self):
        lines = [
            "META ESPECÍFICA 1",
            "Reduzir a violência",
            "Status: Planejado",
            "Descrição do Indicador: Taxa do EVM",
            "Fórmula: (A-B)/A",
            "Carteira de Políticas do MJSP: Política X",
            "Meta do PNSP: Meta PNSP X",
            "Meta do PESP: Meta PESP X",
            "Periodicidade: Anual | Fonte/Ano: SINESP/SENASP/2025",
        ]
        sections = extract_meta_especifica_sections(lines)
        self.assertEqual(len(sections), 1)
        self.assertEqual(sections[0]["periodicidade"], "Anual")
        self.assertEqual(sections[0]["fonte_ano"], "SINESP/SENASP/2025")

    def test_periodicidade_com_valor_referencia_fonte_inline_variacao_nova(self):
        lines = [
            "META ESPECÍFICA 1",
            "Reduzir a violência",
            "Status: Planejado",
            "Descrição do Indicador: Taxa do EVM",
            "Fórmula: (A-B)/A",
            "Carteira de Políticas do MJSP: Política X",
            "Meta do PNSP: Meta PNSP X",
            "Meta do PESP: Meta PESP X",
            "Periodicidade: Anual | Valor de Referência/Fonte: Estupro: 1.956/SINESP/SENASP/2025",
        ]
        sections = extract_meta_especifica_sections(lines)
        self.assertEqual(len(sections), 1)
        self.assertEqual(sections[0]["periodicidade"], "Anual")
        self.assertEqual(
            sections[0]["fonte_ano"],
            "Estupro: 1.956/SINESP/SENASP/2025",
        )


if __name__ == "__main__":
    unittest.main()
