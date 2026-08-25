import sys
import os
import math
from datetime import datetime

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QPixmap

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QTextEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QGridLayout,
    QHBoxLayout,
    QVBoxLayout,
    QGroupBox,
    QTabWidget,
    QCheckBox,
    QListWidget,
    QListWidgetItem,
    QProgressBar,
    QDialog,
    QDialogButtonBox,
    QScrollArea,
    QComboBox,
    QFrame,
    QSizePolicy,
)

from docxtpl import DocxTemplate


# ==========================================================
# PASTAS
# ==========================================================

PASTA_MODELOS_LEGADO = os.path.join(
    os.path.expanduser("~"),
    "Documents",
    "Modelo de relatorio"
)


def obter_diretorio_base():
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


PASTA_MODELOS = os.path.join(obter_diretorio_base(), "templates")


class ModalDetalhes(QDialog):
    def __init__(self, titulo, conteudo, parent=None):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setModal(True)
        self.setMinimumSize(600, 500)

        layout = QVBoxLayout()

        label_titulo = QLabel(titulo)
        label_titulo.setStyleSheet("""
            font-size: 16px;
            font-weight: bold;
            padding: 10px;
            background-color: #ecf0f1;
            border-radius: 5px;
        """)
        layout.addWidget(label_titulo)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setText(conteudo)
        layout.addWidget(text_edit)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(button_box)

        self.setLayout(layout)

class GerenciadorRelatoriosEspecificos:
    """Gerencia a criação e geração de relatórios específicos (barramento e parcelamento)"""

    def __init__(self, parent):
        self.parent = parent

    def obter_dados_comuns(self):
        """Obtém dados comuns de todas as abas"""
        dados = {
            'processo': self.parent.inputs.get('processo', QLineEdit()).text() or '',
            'processo_sei': self.parent.inputs.get('processo', QLineEdit()).text() or '',
            'processo_ina': self.parent.inputs.get('processo', QLineEdit()).text() or '',
            'imovel': self.parent.inputs.get('imovel', QLineEdit()).text() or '',
            'municipio': self.parent.inputs.get('municipio', QLineEdit()).text() or '',
            'uf': self.parent.inputs.get('uf', QLineEdit()).text() or 'GO',
            'car': self.parent.inputs.get('car', QLineEdit()).text() or '',
            'proprietario': self.parent.inputs.get('proprietario', QLineEdit()).text() or '',
            'cpf': self.parent.inputs.get('cpf', QLineEdit()).text() or '',
            'coordenadas': self.parent.inputs.get('coordenadas', QLineEdit()).text() or '',
            'os': self.parent.inputs.get('os', QLineEdit()).text() or '',
            'alertas': self.parent.inputs.get('alertas', QLineEdit()).text() or '',
            'observacoes': self.parent.obs.toPlainText() if hasattr(self.parent, 'obs') else '',
            'intervalo_supressao': self.parent.inputs.get('intervalo_supressao', QLineEdit()).text() or '',
        }
        return dados

    def _numero_extenso(self, valor):
        """Converte número para extenso (ex: 40100.00 -> quarenta mil e cem reais)"""
        try:
            valor_int = int(valor)
            centavos = int(round((valor - valor_int) * 100))

            unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
            especiais = {10: 'dez', 11: 'onze', 12: 'doze', 13: 'treze', 14: 'quatorze',
                         15: 'quinze', 16: 'dezesseis', 17: 'dezessete', 18: 'dezoito', 19: 'dezenove'}
            dezenas = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta',
                       'sessenta', 'setenta', 'oitenta', 'noventa']
            centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos',
                        'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']

            def converter_ate_999(n):
                if n == 0:
                    return ''
                if n == 100:
                    return 'cem'
                texto = ''
                if n >= 100:
                    texto += centenas[n // 100]
                    n %= 100
                    if n > 0:
                        texto += ' e '
                if n >= 20:
                    texto += dezenas[n // 10]
                    n %= 10
                    if n > 0:
                        texto += ' e ' + unidades[n]
                elif 10 <= n <= 19:
                    texto += especiais[n]
                elif n > 0:
                    texto += unidades[n]
                return texto

            if valor_int == 0:
                texto_inteiro = 'zero'
            elif valor_int < 1000:
                texto_inteiro = converter_ate_999(valor_int)
            elif valor_int < 1000000:
                milhares = valor_int // 1000
                resto = valor_int % 1000
                texto_inteiro = converter_ate_999(milhares) + ' mil'
                if resto > 0:
                    if resto < 100:
                        texto_inteiro += ' e '
                    else:
                        texto_inteiro += ' '
                    texto_inteiro += converter_ate_999(resto)
            else:
                milhoes = valor_int // 1000000
                resto = valor_int % 1000000
                texto_inteiro = converter_ate_999(milhoes) + ' milhões'
                if resto > 0:
                    if resto < 100:
                        texto_inteiro += ' e '
                    else:
                        texto_inteiro += ' '
                    if resto < 1000:
                        texto_inteiro += converter_ate_999(resto)
                    else:
                        milhares = resto // 1000
                        resto_mil = resto % 1000
                        texto_inteiro += converter_ate_999(milhares) + ' mil'
                        if resto_mil > 0:
                            if resto_mil < 100:
                                texto_inteiro += ' e '
                            else:
                                texto_inteiro += ' '
                            texto_inteiro += converter_ate_999(resto_mil)

            texto = texto_inteiro + ' reais'
            if centavos > 0:
                if centavos == 1:
                    texto += ' e um centavo'
                else:
                    texto += f' e {converter_ate_999(centavos)} centavos'
            return texto
        except:
            return f"{valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

    def _formatar_moeda_br(self, valor):
        """Formata valor para moeda brasileira"""
        try:
            if isinstance(valor, str):
                # Remove caracteres não numéricos
                valor_limpo = valor.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.').strip()
                if valor_limpo:
                    valor = float(valor_limpo)
                else:
                    return "R$ 0,00"
            valor = float(valor)
            return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
        except (ValueError, TypeError):
            return f"R$ {valor}"
        
    def gerar_relatorio_barramento(self, dados_especificos):
        """Gera o relatório de barramento completo"""
        dados = self.obter_dados_comuns()
        dados.update(dados_especificos)

        # Valores padrão - COM VERIFICAÇÃO DE STRING VAZIA
        try:
            if not dados.get('valor_ai_66') or str(dados.get('valor_ai_66')).strip() == '':
                dados['valor_ai_66'] = 3000.00
            else:
                dados['valor_ai_66'] = float(str(dados.get('valor_ai_66')).replace(',', '.'))
        except:
            dados['valor_ai_66'] = 3000.00

        try:
            if not dados.get('valor_ai_43') or str(dados.get('valor_ai_43')).strip() == '':
                dados['valor_ai_43'] = 5012.50
            else:
                dados['valor_ai_43'] = float(str(dados.get('valor_ai_43')).replace(',', '.'))
        except:
            dados['valor_ai_43'] = 5012.50

        try:
            if not dados.get('valor_recursos_hidricos') or str(dados.get('valor_recursos_hidricos')).strip() == '':
                dados['valor_recursos_hidricos'] = 902.25
            else:
                dados['valor_recursos_hidricos'] = float(str(dados.get('valor_recursos_hidricos')).replace(',', '.'))
        except:
            dados['valor_recursos_hidricos'] = 902.25

        # Calcula total dos autos
        total_autos = float(dados.get('valor_ai_66', 0)) + float(dados.get('valor_ai_43', 0)) + float(
            dados.get('valor_recursos_hidricos', 0))

        # Área de supressão arredondada - COM VERIFICAÇÃO
        try:
            area_supressao = float(str(dados.get('area_supressao', '0')).replace(',', '.'))
            area_arredondada = math.ceil(area_supressao)
            dados['area_supressao_arredondada'] = str(area_arredondada)
        except:
            dados['area_supressao_arredondada'] = '1'

        relatorio = f"""RELATÓRIO DE FISCALIZAÇÃO - BARRAMENTO

Em cumprimento a Ordem de Serviço SEMAD/GO de nº {dados.get('os', 'NÃO INFORMADO')}, no dia {dados.get('data_ocorrencia', 'NÃO INFORMADO')}, esta equipe de fiscalização deslocou-se até o município de {dados.get('municipio', 'NÃO INFORMADO')} para averiguar as informações prestadas na {dados.get('manifestacao', 'NÃO INFORMADO')}, cuja descrição: {dados.get('fato_denunciado', 'NÃO INFORMADO')}

Conforme informado no INÃ PR:
"{dados.get('texto_ina_pr', 'NÃO INFORMADO')}"

Para cumprir com os objetivos estabelecidos, foram empregadas diversas metodologias de fiscalização, incluindo análise de imagens de satélite, observação in loco e documental, visando uma apuração detalhada e precisa da situação.

Tipo da Ação:
( ) Análise/Fiscalização Processual
( ) Fiscalização (remota) – §4° do Art.36 da Lei Estadual 20.694/2019
( ) Fiscalização in loco
( ) Fiscalização em atividade sem licença.

Motivação:
( ) Acompanhamento
( ) Análise Complementar (juntada de documento após fiscalização)
( ) Análise de Atendimento de Notificação
( ) Denúncia/Ouvidoria/INÃ: {dados.get('manifestacao', 'NÃO INFORMADO')} / {dados.get('relatorio_ina', 'NÃO INFORMADO')}
( ) Análise de Processos solicitados por órgãos externos – Processo:
( ) LAI - Lei de Acesso à Informação – Processo:

Localização da Atividade:
() Zona Urbana: {dados.get('imovel', 'NÃO INFORMADO')}, no município {dados.get('municipio', 'NÃO INFORMADO')}, no entorno da Coordenada Geográfica SIRGAS 2000: {dados.get('coordenadas', 'NÃO INFORMADO')}
() Zona Rural: {dados.get('imovel', 'NÃO INFORMADO')}, no município {dados.get('municipio', 'NÃO INFORMADO')}, no entorno da Coordenada Geográfica SIRGAS 2000: {dados.get('coordenadas', 'NÃO INFORMADO')}

Descrição da atividade fiscalizatória
HISTÓRICO PROCESSUAL:
SGA: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
IPÊ: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
SEI: ( ) Nada consta ( ) Não se aplica ( X ) Processo(s)/Cadastro(s): {dados.get('processo_sei', 'NÃO INFORMADO')}
INÃ: ( ) Nada consta ( ) Não se aplica ( X ) Processo(s)/Cadastro(s): {dados.get('processo_ina', 'NÃO INFORMADO')}
WebOutorga / Veredas: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
SIGA: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
SICAR: ( ) Nada consta ( ) Não se aplica ( X ) Processo(s)/Cadastro(s): {dados.get('car', 'NÃO INFORMADO')}
SEISB: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):

CONSTATAÇÕES
Constatações Técnicas da Fiscalização
( ) A equipe foi recebida por: {dados.get('proprietario', 'NÃO INFORMADO')};
( ) Ninguém foi encontrado no local.

Observações:
Área localizada na zona rural do município de {dados.get('municipio', 'NÃO INFORMADO')};
Na propriedade, foi constatada a realização de obras em uma represa localizada no entorno da Coordenada Geográfica SIRGAS 2000: {dados.get('coordenadas', 'NÃO INFORMADO')}. O {dados.get('proprietario', 'NÃO INFORMADO')} declarou tratar-se de uma represa antiga e que, juntamente com o proprietário da área confrontante (lado oposto), realizou apenas a revitalização da área, há aproximadamente seis meses.

No momento da fiscalização, constatou-se que o reservatório foi objeto de intervenções recentes, sendo possível observar sinais de movimentação de terra e taludes ainda sem cobertura vegetal, que no maciço do talude houve plantio recente de grama. O reservatório possui 02 (dois) extravasores nas ombreiras direita e esquerda e descarga de fundo, por meio de 02 (dois) tubos de PVC no centro do barramento. No momento da fiscalização, observou-se o vertimento de água pela descarga de fundo e por ambas as ombreiras.

Na mesma data, tendo em vista que a denúncia tratava-se de várias intervenções ao longo do curso hídrico, após fiscalização no proprietário do outro lado da barragem, o mesmo informou que nessa localidade a obra foi efetuada pelo {'proprietario'}, o qual não se responsabilizou no momento da fiscalização. Mas foi observado in loco que o {dados.get('proprietario', 'NÃO INFORMADO')} é o principal usuário, sendo percebido uma embarcação e algumas tambores para ração, os quais são utilizados para alimentar os peixes colocados no barramento.

Foram realizadas a coleta de dados in loco e, posteriormente, a consolidação e análise das informações obtidas.

Foi realizado sobrevoo da área por meio de Aeronave Remotamente Pilotada - RPA, modelo DJI Mavic 3, no dia {dados.get('data_sobrevoo', 'NÃO INFORMADO')}, com intuito de coletar imagens aéreas das áreas em questão, possibilitando a confecção de ortofoto das áreas de interesse por meio do software WebODM e a realização de análises de geoprocessamento por meio do software QGIS. Foi elaborado mapa digital de situação contendo as áreas de interesse.

Com objetivo de subsidiar a atividade de fiscalização ambiental foi elaborado o seguinte produto cartográfico:
- Mapa_Geral - {dados.get('mapa_geral', 'NÃO INFORMADO')}
- Mapa_Temporal - {dados.get('mapa_temporal', 'NÃO INFORMADO')}

Para a análise da área foram utilizadas imagens de janeiro, maio, e junho de 2025, da constelação Planet com resolução espacial de 4,77 m, disponibilizadas pela Iniciativa Internacional sobre Clima e Florestas da Noruega-NICFI (https://www.planet.com/nicfi/). Foram utilizados dados disponibilizados no Sistema de Informações Geográficas do Estado de Goiás – SIGA. Todas as informações geoespaciais estão georreferenciadas ao Sistema de Coordenadas UTM 22S, Datum SIRGAS2000, utilizando o software livre de Geoprocessamento QGIS.

Constatado:
1 reservatório/barramento:
- Lâmina d'água: {dados.get('area_barramento', 'NÃO INFORMADO')} hectares;
- Supressão/Dano: {dados.get('area_supressao', 'NÃO INFORMADO')} hectares {dados.get('intervalo_supressao', 'NÃO INFORMADO')}.

Diante do exposto, conclui-se que a operação do reservatório em questão encontra-se desprovida de licença ambiental válida, bem como inexistente outorga de direito de uso de recursos hídricos para a acumulação hídrica, em desacordo com a legislação ambiental vigente. Ressalta-se que os danos à vegetação nativa constatados na área decorrem diretamente da implantação do barramento. A vegetação existente foi suprimida durante a execução da obra e, posteriormente, danificada pelo acúmulo permanente de água, o que ocasionou a morte da cobertura vegetal remanescente. Ademais, a formação do reservatório alterou as condições naturais do ambiente, impedindo a regeneração natural da vegetação nas áreas atingidas, especialmente na Área de Preservação Permanente (APP).

Em razão das irregularidades constatadas, foram adotadas as medidas legais cabíveis, nos termos da normativa aplicável, visando à regularização da atividade e à proteção dos recursos ambientais. O presente relatório consolida os levantamentos e análises realizadas, subsidiando os encaminhamentos administrativos pertinentes no âmbito desta Secretaria.

Sendo assim, foram lavrados em nome de {dados.get('proprietario', 'NÃO INFORMADO')}, CPF: {dados.get('cpf', 'NÃO INFORMADO')}, os Autos de Infração nº {dados.get('auto_barramento', 'NÃO INFORMADO')} e Termo de Embargo {dados.get('embargo_barramento', 'NÃO INFORMADO')}, e devidamente enviados por meio de carta registrada com aviso de recebimento, para o endereço de correspondência informado pelo {dados.get('proprietario', 'NÃO INFORMADO')} no momento da fiscalização.

LEGISLAÇÃO PERTINENTE
- Decreto Federal nº 6.514/2008;
- Decreto Estadual nº 9.710/2020;
- Decreto Estadual nº 10.371/2023;
- Lei Federal nº 9.605/1998;
- Lei Estadual nº 20.694/2019;
- Lei Estadual nº 18.102/2013;
- Lei Estadual nº 13.123/1997.

INFORMAÇÕES COMPLEMENTARES:

DECRETO Nº 10.371, DE 20 DE DEZEMBRO DE 2023
Altera o Decreto nº 9.710, de 3 de setembro de 2020, que regulamenta, no âmbito do Poder Executivo estadual, a Lei nº 20.694, de 26 de dezembro de 2019, que dispõe sobre as normas gerais para o licenciamento ambiental no Estado de Goiás e dá outras providências.

ANEXO ÚNICO (DECRETO Nº 9.710, DE 3 DE SETEMBRO DE 2020)
DIVISÃO "F": OBRAS CIVIS
Grupo F2: barragens, diques e canais.
F2.1 - Reservatórios e diques para captação de água de chuva ou derivada, fora de APP e leito de rio perene ou intermitente - Lâmina de água do reservatório (ha) - Micro ≥ 1 < 5
F2.2 - Reservatórios/barragens e diques em curso de água para abastecimento humano, dessedentação animal, irrigação, fins paisagísticos* e composição urbana, lazer, turismo e aquicultura sem remoção de pessoas. * para fins paisagísticos e composição urbana, lazer ou turismo, somente com decreto do Chefe do Poder Executivo estadual ou federal; e ** as barragens instaladas depois de 27 de dezembro de 2019 com área do reservatório menor do que 1,2 ha e para os fins descritos acima deverão ser enquadradas na tipologia F2.6 - Lâmina de água do reservatório (ha) Micro ≥ 1,2** < 5
F2.6 - Reservatórios/barragens e diques em curso de água com lâmina d'água entre 0,1 e 1,2 hectare para abastecimento humano, dessedentação animal, irrigação, fins paisagísticos* e composição urbana, lazer, turismo e aquicultura sem remoção de pessoas. * para fins paisagísticos e outros usos não previstos em lei, é necessário o decreto de utilidade pública - Lâmina de água do reservatório (ha) - Micro ≥ 0,1 < 1,2.

LEI COMPLEMENTAR Nº 140, DE 8 DE DEZEMBRO DE 2011
No artigo 3º da Resolução CEMAm 259/2024, especifica quais os parâmetros e requisitos o município deverá atender para o credenciamento para licenciar as atividades que estão definidas no anexo único da mesma resolução.
Hidrolândia - Res. CEMAm n°174, de 18 de Outubro de 2022 - Nível 2.

1. Que atividade(s) é(são) realizada(s) no local?
Barramento em área de APP e acumulação de água sem autorização do órgão ambiental competente;

2. A(s) atividade(s) é(são) utilizadora(s) de recursos ambientais, consideradas efetiva ou potencialmente poluidoras, ou capazes de, sob qualquer forma, causar degradação ambiental?
Sim;

3. A(s) atividade(s) é(são) licenciada(s)/autorizada(s)?
Não;

4. Quem é(são) o(s) responsável(is) pela(s) atividade(s)?
{dados.get('proprietario', 'NÃO INFORMADO')}, CPF: {dados.get('cpf', 'NÃO INFORMADO')};

5. Qual o endereço e coordenadas geográficas do local onde é(são) realizada(s) a(s) atividade(s)?
{dados.get('imovel', 'NÃO INFORMADO')}, no município de {dados.get('municipio', 'NÃO INFORMADO')}, no entorno da Coordenada Geográfica SIRGAS 2000: {dados.get('coordenadas', 'NÃO INFORMADO')};

6. Quem é(são) o(s) proprietário(s) da(s) área(s) onde é(são realizadas(s) a(s) atividade(s)?
{dados.get('proprietario', 'NÃO INFORMADO')}, CPF: {dados.get('cpf', 'NÃO INFORMADO')};

7. A(s) atividade(s) é(são) desenvolvida(s) em áreas protegidas (APP, Reserva Legal e/ou Unidade de Conservação)?
Sim;

8. A(s) atividade(s) causou(ram) ou está(ão) causando degradação ambiental? Se sim, em que consiste a degradação ambiental?
Sim. Consiste na operação de barragem em área de APP, sem autorização do órgão ambiental competente;

9. Que medidas foram ou devem ser adotadas para cessar a degradação ambiental?
Foi realizado o embargo da área e das atividades nelas realizadas;

10. A degradação ambiental comporta recuperação?
Sim, a degradação ambiental pode ser passível de recuperação, embora a viabilidade e o tempo necessário para a recuperação possam variar dependendo da extensão e gravidade da degradação, bem como das características do ecossistema afetado;

11. Que medidas foram ou devem ser adotadas para promover a recuperação ambiental?
As medidas técnicas e ambientais serão estabelecidas no âmbito da licença ambiental do empreendimento ou do termo de compromisso específico;

DA VALORAÇÃO

{dados.get('auto_artigo_66', 'NÃO INFORMADO')}
O artigo 66 do Decreto nº 6.514/2008 estabelece a valoração de Multa de R$ 500,00 (quinhentos reais) a R$ 10.000.000,00 (dez milhões de reais):
Art. 66. Construir, reformar, ampliar, instalar ou fazer funcionar estabelecimentos, atividades, obras ou serviços utilizadores de recursos ambientais, considerados efetiva ou potencialmente poluidores, sem licença ou autorização dos órgãos ambientais competentes, em desacordo com a licença obtida ou contrariando as normas legais e regulamentos pertinentes:
Multa de R$ 500,00 (quinhentos reais) a R$ 10.000.000,00 (dez milhões de reais).

Para fixação do valor referente a infração foram utilizados os critérios definidos na ORIENTAÇÃO NORMATIVA SEMAD Nº 01/2024, considerando o nível de gravidade da infração e identificação da capacidade econômica da seguinte maneira:
- Motivo da Infração: Obtenção de vantagem pecuniária (15)
- Consequência para o meio ambiente: Moderada (30)
- Consequência para a saúde pública ou para a socioeconomia da área de abrangência do fato: Não Houve (0)
Somatório dos valores desta etapa: (45) - Nível B
Situação econômica: Não foi possível aferir a situação econômica do infrator. Por conta disso, foi utilizado Situação econômica - Receita Mensal - Pessoa Física - Receita bruta: mensal de até 1 salário mínimo (Faixa A) = Nível B: Mínimo + 0,025% até 0,8% do teto.

VALORAÇÃO: R$ 500,00 + (0,025% X R$ 10.000.000,00) = {self._formatar_moeda_br(dados.get('valor_ai_66', 3000))} ({self._numero_extenso(dados.get('valor_ai_66', 3000))})

Considerando-se que foi utilizado a classe de menor valor para a situação econômica do infrator, que pode não ser condizente com a realidade, sugere-se na Autocomposição/julgamento, a aferição e reenquadramento se for o caso conforme Art. 8º, §2º ON 01/2024 SEMAD.
Art. 8º Em se tratando de pessoa física, a situação econômica do infrator será determinada pelos critérios estabelecidos no
Quadro 3 do Anexo único, mediante a classificação em faixas definidas conforme receita bruta anual do infrator, assim estabelecidas:
I - receita bruta mensal de até 1 salário mínimo;
II - receita bruta mensal, superior ao limite anterior até 3 salários mínimos;
III - receita bruta mensal, superior ao limite anterior até 10 salários mínimos;
IV - receita bruta mensal, superior ao limite anterior até 30 salários mínimos;
V - receita bruta mensal, superior ao limite anterior até 45 salários mínimos; e
VI - receita bruta mensal, superior ao limite anterior.
§ 1º Em se tratando de pessoa física serão considerados os rendimentos indicados em qualquer documento válido para comprovação de renda.
§ 2º A autoridade julgadora competente bem como os facilitadores em sede de audiências de autocomposição deverão rever o enquadramento do infrator quanto a sua situação econômica, caso conste no relatório de fiscalização que esta não tenha sido possível aferir.

Lei Estadual 13.123, DE 16 DE JULHO DE 1997.
SEÇÃO II DAS INFRAÇÕES E PENALIDADES
Art. 13. Constitui infração às normas de utilização de recursos hídricos superficiais e subterrâneos:
I - derivar ou utilizar dos recursos hídricos para qualquer finalidade, sem a respectiva outorga de direito de uso;
II - iniciar a implantação ou implantar empreendimento relacionado com a derivação ou utilização de recursos hídricos, superficiais e/ou subterrâneos, que implique alterações no regime, quantidade e qualidade dos mesmos, sem autorização dos órgãos ou entidades competentes;
III - deixar expirar o prazo de validade das outorgas sem solicitar a devida prorrogação ou revalidação;
IV - utilizar-se dos recursos hídricos ou executar obras ou serviços relacionados com os mesmos em desacordo com as condições estabelecidas na outorga;
Art. 14. Por infração a qualquer disposição legal ou regulamentar referente a execução de obras e serviços hidráulicos, derivação e utilização de recursos hídricos de domínio ou administração do Estado de Goiás, ou pelo não atendimento das solicitações feitas, o infrator, a critério da autoridade competente, ficará sujeito às seguintes penalidades, independentemente da sua ordem de enumeração:
II - multa, simples ou diária, proporcional à gravidade da infração, de R$ 90,00 (noventa reais) a R$ 90.000,00 (noventa mil reais), corrigidos pela UFIR;
Art. 15. As infrações às disposições desta lei às normas dela decorrentes serão, a critério da autoridade impositora, classificadas em leves, graves, gravíssimas, levando em conta:
I - as circunstâncias atenuantes e agravantes;
II - os antecedentes do infrator.
§ 1o As multas simples ou diárias, a critério da autoridade aplicadora, ficam estabelecidas dentro das seguintes faixas:
b) acima de R$ 900,00 (novecentos reais) até 9.000,00 (nove mil reais), nas infrações graves;

Resolução CERHi nº 66, de 26 de janeiro de 2024, artigo 3°, inciso V:
Art. 3°. Estão sujeitos à outorga:
V – as acumulações de água em corpos hídricos;
Para fixação do valor referente a infração foi utilizado os critérios definidos nos quadros 1 e 3 da ORIENTAÇÃO NORMATIVA SEMAD Nº 1/2024 - GAB- 06281 e por não contemplar as classificações em infrações leves, graves e gravíssimas referente ao funcionamento/operação de barragem sem a respectiva outorga de direito de uso, foi utilizado por analogia a RESOLUÇÃO Nº 24, DE 04 DE MAIO DE 2020 da Agência Nacional das Águas - ANA, considerando o inciso I, do Artigo 18, o nível de gravidade da infração e identificação da capacidade econômica da seguinte maneira:
Art. 18. São consideradas infrações graves:
I – derivar ou utilizar recursos hídricos para qualquer finalidade, sem a respectiva outorga de direito de uso;

Para fixação do valor referente a infração foram utilizados os critérios definidos na ORIENTAÇÃO NORMATIVA SEMAD Nº 01/2024 (Define parâmetros para a fixação das multas abertas, para a aplicação de sanções e medidas administrativas cautelares no âmbito da apuração de infrações ambientais, bem como critérios para o agravamento e a atenuação das sanções administrativas decorrentes de infrações ambientais), considerando o nível de gravidade da infração e identificação da capacidade econômica. Gravidade da Infração:
- Motivo da Infração: Obtenção de vantagem pecuniária (15)
- Consequência para o meio ambiente: Fraca (20)
- Consequência para a saúde pública ou para a socioeconomia da área de abrangência do fato: Não Houve (0)
Somatório dos valores desta etapa: (35) - Nível B
Situação econômica: Não foi possível aferir a situação econômica do infrator. Por conta disso, foi utilizado Situação econômica - Receita Mensal - Pessoa Física - Receita bruta: mensal de até 1 salário mínimo (Faixa A) = Nível B: Mínimo + 0,025% até 0,8% do teto.

VALORAÇÃO: (R$ 900,00 + (0,025% X R$ 9.000,00)), totalizando o valor de {self._formatar_moeda_br(dados.get('valor_recursos_hidricos', 902.25))} ({self._numero_extenso(dados.get('valor_recursos_hidricos', 902.25))}).

Considerando-se que foi utilizado a classe de menor valor para a situação econômica do infrator, que pode não ser condizente com a realidade, sugere-se na Autocomposição/julgamento, a aferição e reenquadramento se for o caso conforme Art. 8º, §2º ON 01/2024 SEMAD.

{dados.get('auto_artigo_43', 'NÃO INFORMADO')}
O artigo 43 do Decreto nº 6.514/2008 estabelece a valoração de Multa de R$ 5.000,00 (cinco mil reais) a R$ 50.000,00 (cinquenta mil reais), por hectare ou fração. :
Art. 43. Destruir ou danificar florestas ou demais formas de vegetação natural ou utilizá-las com infringência das normas de proteção em área considerada de preservação permanente, sem autorização do órgão competente, quando exigível, ou em desacordo com a obtida: 
Para fixação do valor referente a infração foram utilizados os critérios definidos na ORIENTAÇÃO NORMATIVA SEMAD Nº 01/2024 (Define parâmetros para a fixação das multas abertas, para a aplicação de sanções e medidas administrativas cautelares no âmbito da apuração de infrações ambientais, bem como critérios para o agravamento e a atenuação das sanções administrativas decorrentes de infrações ambientais), considerando o nível de gravidade da infração e identificação da capacidade econômica da seguinte maneira:
- Motivo da Infração: Obtenção de vantagem pecuniária (15)
- Consequência para o meio ambiente: Moderada (30)
- Consequência para a saúde pública ou para a socioeconomia da área de abrangência do fato: Não Houve (0)
Somatório dos valores desta etapa: (45) - Nível B
Situação econômica: Não foi possível aferir a situação econômica do infrator. Por conta disso, foi utilizado Situação econômica - Receita Mensal - Pessoa Física - Receita bruta: mensal de até 1 salário mínimo (Faixa A) = Nível B: Mínimo + 0,025% até 0,8% do teto.

R$ 5.000,00 + (0,025% X R$ 50.000,00) = {self._formatar_moeda_br(dados.get('valor_ai_43', 5012.50))} ({self._numero_extenso(dados.get('valor_ai_43', 5012.50))})

Valoração: {dados.get('area_supressao', 'NÃO INFORMADO')} hectares x {self._formatar_moeda_br(dados.get('valor_ai_43', 5012.50))} = {self._formatar_moeda_br(dados.get('valor_ai_43', 5012.50) * float(dados.get('area_supressao_arredondada', 1)))} ({self._numero_extenso(dados.get('valor_ai_43', 5012.50) * float(dados.get('area_supressao_arredondada', 1)))})

{'=' * 60}
SOMA TOTAL DOS AUTOS

A soma dos autos totaliza o valor de {self._formatar_moeda_br(total_autos)} ({self._numero_extenso(total_autos)}).

Considerando-se que foi utilizado a classe de menor valor para a situação econômica do infrator, que pode não ser condizente com a realidade, sugere-se na Autocomposição/julgamento, a aferição e reenquadramento se for o caso conforme Art. 8º, §2º ON 01/2024 SEMAD.

{'=' * 60}
{dados.get('municipio', 'NÃO INFORMADO')}, {dados.get('data_ocorrencia', 'NÃO INFORMADO')}
_____________________________________
Assinatura do Fiscal
"""
        return relatorio
  

class GeradorRelatoriosThread(QThread):
    progresso = Signal(int)
    status = Signal(str)
    concluido = Signal(list)
    erro = Signal(str)

    def __init__(self, modelos_selecionados, contexto, pasta_destino):
        super().__init__()
        self.modelos_selecionados = modelos_selecionados
        self.contexto = contexto
        self.pasta_destino = pasta_destino
        self.arquivos_gerados = []

    def run(self):
        try:
            total = len(self.modelos_selecionados)
            for i, (nome_modelo, caminho_modelo, tipo) in enumerate(self.modelos_selecionados):
                self.status.emit(f"Gerando relatório {tipo.upper()}...")
                
                if not os.path.exists(caminho_modelo):
                    raise FileNotFoundError(
                        f"Modelo não encontrado:\n{caminho_modelo}"
                    )

                doc = DocxTemplate(caminho_modelo)
                doc.render(self.contexto)
                
                nome_arquivo = (
                    f"{tipo}_"
                    f"{self.contexto['imovel']}_"
                    f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                )
                
                nome_arquivo = "".join(
                    c for c in nome_arquivo
                    if c.isalnum() or c in '._-'
                )
                
                caminho_saida = os.path.join(
                    self.pasta_destino,
                    nome_arquivo
                )
                
                doc.save(caminho_saida)
                self.arquivos_gerados.append(caminho_saida)
                
                progresso = int((i + 1) / total * 100)
                self.progresso.emit(progresso)
                
            self.concluido.emit(self.arquivos_gerados)
            
        except Exception as e:
            self.erro.emit(str(e))


# ==========================================================
# CLASSE PRINCIPAL
# ==========================================================

class SistemaRelatorios(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Apoio à Minuta e Relatórios Ambientais - SEMAD")
        self.setGeometry(100, 50, 1300, 850)

        self.inputs = {}
        self.areas_inputs = {}
        self.artigos_inputs = {}

        # ==========================================================
        # INICIALIZAR DICIONÁRIOS DOS CAMPOS DE RELATÓRIOS ESPECÍFICOS
        # ==========================================================
        self.campos_barramento = {}
        self.campos_parcelamento = {}

        # Inicializar checkboxes como None para evitar erros
        self.check_app = None
        self.check_rl = None
        self.check_fora = None
        self.check_dano = None
        self.check_art48 = None
        self.check_art66 = None
        self.check_art79 = None
        self.check_uc = None
        self.check_uc_artigos = None

        # Checkboxes para os novos relatórios
        self.check_relatorio_ina = None
        self.check_despacho = None
        self.check_minuta = None
        self.check_autodenuncia = None

        self.gerenciador = GerenciadorRelatoriosEspecificos(self)
        self.aplicar_tema_goias()
        self.init_interface()
        self.adicionar_tooltips()
        self.conectar_checkboxes_artigos()

    def conectar_checkboxes_artigos(self):
        """Conecta os checkboxes dos artigos para atualizar a pré-visualização"""

        # Art. 48
        if hasattr(self, 'check_art48') and self.check_art48:
            self.check_art48.stateChanged.connect(self.atualizar_preview)

        # Art. 66
        if hasattr(self, 'check_art66') and self.check_art66:
            self.check_art66.stateChanged.connect(self.atualizar_preview)

        # Art. 79
        if hasattr(self, 'check_art79') and self.check_art79:
            self.check_art79.stateChanged.connect(self.atualizar_preview)

    def preencher_texto_artigo(self, state, tipo_artigo):
        """Preenche os campos do artigo quando o checkbox é marcado"""
        if state == Qt.Checked:
            # Obtém os dados do artigo
            artigo_data = self.artigos_inputs.get(tipo_artigo, {})
            area_input = artigo_data.get("area")
            auto_input = artigo_data.get("auto")
            embargo_input = artigo_data.get("embargo")
            valor_input = artigo_data.get("valor")

            # Se não houver área preenchida, usa um valor padrão
            if area_input and not area_input.text().strip():
                area_input.setText("1")

            # Se não houver valor calculado, usa um placeholder
            if valor_input and not valor_input.text().strip():
                if tipo_artigo == "art48":
                    # Art. 48 - multa simples
                    area_text = area_input.text().strip() if area_input else "1"
                    try:
                        area_float = float(area_text.replace(',', '.'))
                        area_arredondada = math.ceil(area_float)
                        valor_total = area_arredondada * 5000.00
                        valor_input.setText(
                            f"R$ {valor_total:,.2f}".replace(",", "v").replace(".", ",").replace("v", "."))
                    except:
                        valor_input.setText("R$ 5.000,00")
                else:
                    # Art. 66 e 79 - dosimetria (valor estimado)
                    valor_input.setText("R$ 5.000,00 a R$ 50.000,00")

            # Atualiza a pré-visualização
            self.atualizar_preview()

    def on_uc_changed(self, state):
        """Quando o checkbox UC é alterado, recalcula todas as multas sem alterar os textos"""
        for chave in ["app", "rl", "fora", "dano"]:
            if chave in self.areas_inputs:
                area_widget = self.areas_inputs[chave].get("area")
                valor_widget = self.areas_inputs[chave].get("valor")

                if area_widget and area_widget.text().strip():
                    auto_original = ""
                    embargo_original = ""
                    auto_widget = self.areas_inputs[chave].get("auto")
                    embargo_widget = self.areas_inputs[chave].get("embargo")

                    if auto_widget and hasattr(auto_widget, 'toPlainText'):
                        auto_original = auto_widget.toPlainText()
                    if embargo_widget and hasattr(embargo_widget, 'toPlainText'):
                        embargo_original = embargo_widget.toPlainText()

                    if chave == "app":
                        self.calcular_valor_multa_app(chave, area_widget, valor_widget)
                    else:
                        self.calcular_multa_simples(chave, area_widget, valor_widget)

                    if auto_original and auto_widget:
                        auto_widget.setText(auto_original)
                    if embargo_original and embargo_widget:
                        embargo_widget.setText(embargo_original)

        self.atualizar_preview()

    
    def on_uc_changed(self, state):
        """Quando o checkbox UC é alterado, recalcula todas as multas sem alterar os textos"""
        # Verifica se já existem áreas com valores calculados
        for chave in ["app", "rl", "fora", "dano"]:
            if chave in self.areas_inputs:
                area_widget = self.areas_inputs[chave].get("area")
                valor_widget = self.areas_inputs[chave].get("valor")
                
                # Se tem área preenchida, recalcula
                if area_widget and area_widget.text().strip():
                    # Salvar os textos originais dos campos de Auto e Embargo
                    auto_original = ""
                    embargo_original = ""
                    auto_widget = self.areas_inputs[chave].get("auto")
                    embargo_widget = self.areas_inputs[chave].get("embargo")
                    
                    if auto_widget and hasattr(auto_widget, 'toPlainText'):
                        auto_original = auto_widget.toPlainText()
                    if embargo_widget and hasattr(embargo_widget, 'toPlainText'):
                        embargo_original = embargo_widget.toPlainText()
                    
                    # Recalcular a multa (isso vai atualizar o valor e pode chamar atualizar_texto_supressao_apos_calculo)
                    if chave == "app":
                        self.calcular_valor_multa_app(chave, area_widget, valor_widget)
                    else:
                        self.calcular_multa_simples(chave, area_widget, valor_widget)
                    
                    # Restaurar os textos originais (se não estiverem vazios)
                    if auto_original and auto_widget:
                        auto_widget.setText(auto_original)
                    if embargo_original and embargo_widget:
                        embargo_widget.setText(embargo_original)

        # ==========================================================
        # RECALCULA OS ARTIGOS DA ABA "AUTOS E EMBARGOS (DEMAIS)"
        # ==========================================================
        # Art. 48 - multa simples (não depende de dosimetria)
        if hasattr(self, 'artigos_inputs') and "art48" in self.artigos_inputs:
            dados_art48 = self.artigos_inputs["art48"]
            area_widget = dados_art48.get("area")
            valor_widget = dados_art48.get("valor")
            if area_widget and valor_widget and area_widget.text().strip():
                self.calcular_multa_artigo_simples(area_widget, valor_widget, 5000.00)

        # Art. 66 e Art. 79 - dosimetria (só recalcula se já houver um valor calculado)
        for chave_artigo, valor_base_artigo, teto_artigo in [
            ("art66", 500.00, 10000000.00),
            ("art79", 10000.00, 10000000.00),
        ]:
            if chave_artigo in self.artigos_inputs:
                dados_artigo = self.artigos_inputs[chave_artigo]
                valor_widget = dados_artigo.get("valor")
                pontuacao_widget = dados_artigo.get("pontuacao")
                if valor_widget and valor_widget.text().strip() and pontuacao_widget and pontuacao_widget.text().strip():
                    self.calcular_multa_com_dosimetria(
                        chave_artigo,
                        dados_artigo.get("area"),
                        valor_widget,
                        dados_artigo.get("motivo"),
                        dados_artigo.get("consequencia"),
                        dados_artigo.get("saude"),
                        pontuacao_widget,
                        dados_artigo.get("tipo_infrator"),
                        dados_artigo.get("faixa_receita"),
                        dados_artigo.get("percentual"),
                        valor_base_artigo,
                        teto_artigo,
                    )

        # Atualiza a pré-visualização
        self.atualizar_preview()
    def atualizar_faixas_receita(self):
        if hasattr(self, 'faixa_receita_combo'):
            self.faixa_receita_combo.clear()
            if self.tipo_infrator_combo.currentText() == "Pessoa Jurídica":
                self.faixa_receita_combo.addItems(self.faixas_pj)
            else:
                self.faixa_receita_combo.addItems(self.faixas_pf)

    def aplicar_tema_goias(self):
        self.setStyleSheet("""
            /* Reset e base */
            QWidget {
                background-color: #f0f2f5;
                font-family: 'Segoe UI', 'Roboto', Arial, sans-serif;
                font-size: 10pt;
                color: #2c3e50;
            }
            
            /* Tabs */
            QTabWidget::pane {
                border: 1px solid #e1e4e8;
                background: white;
                border-radius: 10px;
                top: -1px;
                padding: 10px;
            }
            QTabBar::tab {
                background: transparent;
                color: #586069;
                padding: 12px 28px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: 500;
                font-size: 10pt;
                min-width: 100px;
            }
            QTabBar::tab:selected {
                background: white;
                color: #006b3f;
                border: 1px solid #e1e4e8;
                border-bottom-color: white;
                font-weight: 600;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ffffff, stop:1 #f8f9fa);
            }
            QTabBar::tab:hover:!selected {
                background: #f6f8fa;
                color: #006b3f;
            }
            
            /* GroupBox moderno */
            QGroupBox {
                border: 1px solid #e1e4e8;
                border-radius: 10px;
                margin-top: 16px;
                padding-top: 16px;
                background-color: white;
                font-weight: 500;
                color: #006b3f;
                font-size: 10pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 12px 0 12px;
                background-color: white;
                color: #006b3f;
                font-weight: 600;
                font-size: 11pt;
            }
            
            /* Cards internos */
            QGroupBox[class="area-card"] {
                background-color: #fafbfc;
                border: 1px solid #e8ecf0;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 12px;
            }
            QGroupBox[class="area-card"]::title {
                color: #006b3f;
                font-weight: 600;
                font-size: 10pt;
                padding: 0 8px;
            }
            
            /* Inputs */
            QLineEdit, QTextEdit, QListWidget, QComboBox {
                background-color: #ffffff;
                border: 1.5px solid #d1d5da;
                border-radius: 8px;
                padding: 8px 12px;
                min-height: 30px;
                font-size: 10pt;
                selection-background-color: #006b3f;
                transition: border-color 0.2s;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 2px solid #006b3f;
                background-color: #fefefe;
                box-shadow: 0 0 0 3px rgba(0, 107, 63, 0.1);
            }
            QLineEdit:read-only {
                background-color: #f6f8fa;
                color: #586069;
                border-style: dashed;
            }
            QLineEdit:disabled, QTextEdit:disabled, QComboBox:disabled {
                background-color: #f1f3f4;
                color: #9aa0a6;
                border-color: #e1e4e8;
            }
            
            /* Botões */
            QPushButton {
                background-color: #006b3f;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 22px;
                font-weight: 600;
                font-size: 10pt;
                min-height: 34px;
                transition: all 0.2s;
            }
            QPushButton:hover {
                background-color: #008c52;
                transform: translateY(-1px);
                box-shadow: 0 4px 12px rgba(0, 107, 63, 0.3);
            }
            QPushButton:pressed {
                background-color: #004f2e;
                transform: translateY(0px);
                box-shadow: none;
            }
            QPushButton:disabled {
                background-color: #d1d5da;
                color: #8b949e;
            }
            QPushButton#secondary {
                background-color: #f6f8fa;
                color: #24292e;
                border: 1px solid #d1d5da;
            }
            QPushButton#secondary:hover {
                background-color: #e1e4e8;
                border-color: #b0b8c0;
                box-shadow: none;
            }
            QPushButton#danger {
                background-color: #dc3545;
            }
            QPushButton#danger:hover {
                background-color: #c82333;
            }
            
            /* Checkboxes modernos */
            QCheckBox {
                spacing: 10px;
                padding: 6px 4px;
                color: #24292e;
                font-weight: 500;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border-radius: 4px;
                border: 2px solid #d1d5da;
                background-color: white;
            }
            QCheckBox::indicator:checked {
                background-color: #006b3f;
                border-color: #006b3f;
                image: url(:/icons/check.svg);
            }
            QCheckBox::indicator:hover {
                border-color: #006b3f;
            }
            QCheckBox:disabled {
                color: #9aa0a6;
            }
            QCheckBox:disabled::indicator {
                background-color: #f1f3f4;
                border-color: #e1e4e8;
            }
            
            /* Checkbox UC destaque */
            QCheckBox#uc-check {
                background-color: #fff3cd;
                border-radius: 6px;
                padding: 8px 12px;
                border: 1px solid #ffeeba;
                font-weight: 600;
                color: #856404;
            }
            QCheckBox#uc-check::indicator:checked {
                background-color: #ff6b00;
                border-color: #ff6b00;
            }
            
            /* Progress Bar */
            QProgressBar {
                border: 1px solid #e1e4e8;
                border-radius: 8px;
                text-align: center;
                background-color: #f6f8fa;
                min-height: 28px;
                color: #24292e;
                font-weight: 500;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #006b3f, stop:1 #00a859);
                border-radius: 7px;
            }
            
            /* ScrollArea */
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background: #f6f8fa;
                width: 10px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: #d1d5da;
                border-radius: 5px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background: #006b3f;
            }
            
            /* Labels */
            QLabel {
                color: #24292e;
                padding: 3px;
                font-weight: normal;
            }
            QLabel[class="label-title"] {
                font-weight: 600;
                color: #006b3f;
                font-size: 11pt;
                padding: 8px 0;
            }
            QLabel[class="label-desc"] {
                color: #586069;
                font-style: italic;
                font-size: 9pt;
            }
            QLabel[class="label-obrigatorio"] {
                color: #dc3545;
                font-weight: 600;
                font-size: 9pt;
            }
            
            /* Frames */
            QFrame[class="card-info"] {
                background-color: #e8f5e9;
                border-radius: 8px;
                padding: 12px;
                border-left: 4px solid #006b3f;
            }
            QFrame[class="card-warning"] {
                background-color: #fff3cd;
                border-radius: 8px;
                padding: 12px;
                border-left: 4px solid #ffc107;
            }
            QFrame[class="card-disabled"] {
                background-color: #f1f3f4;
                border-radius: 8px;
                padding: 12px;
                border: 1px dashed #d1d5da;
            }
        """)

    def init_interface(self):
        layout = QVBoxLayout()

        cabecalho = QWidget()
        cabecalho.setStyleSheet("""
            background-color: #006b3f;
            border-radius: 10px;
        """)
        cabecalho_layout = QVBoxLayout(cabecalho)
        cabecalho_layout.setContentsMargins(20, 16, 20, 16)
        cabecalho_layout.setSpacing(4)

        titulo = QLabel("Sistema de Apoio à Minuta e Relatórios Ambientais")
        titulo.setAlignment(Qt.AlignCenter)
        titulo.setStyleSheet("""
            background: transparent;
            color: white;
            font-size: 19px;
            font-weight: 600;
            letter-spacing: 0.3px;
        """)
        cabecalho_layout.addWidget(titulo)

        subtitulo = QLabel("SEMAD")
        subtitulo.setAlignment(Qt.AlignCenter)
        subtitulo.setStyleSheet("""
            background: transparent;
            color: #cdeee0;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: 3px;
        """)
        cabecalho_layout.addWidget(subtitulo)

        layout.addWidget(cabecalho)

        tabs = QTabWidget()
        tabs.addTab(self.criar_aba_dados_gerais(), "📋 Dados Gerais")
        tabs.addTab(self.criar_aba_areas(), "🚫 Autos e Embargos (Supressão de vegetação nativa)")
        tabs.addTab(self.criar_aba_artigos(), "🚫 Autos e Embargos (Demais)")
        tabs.addTab(self.criar_aba_areas_selecionadas(), "📍 Áreas Selecionadas")
        tabs.addTab(self.criar_aba_relatorios_especificos(), "📄 Dados para o relátorio")
        tabs.addTab(self.criar_aba_novos_modelos(), "📄 Relatórios")
        tabs.addTab(self.criar_aba_creditos(), "👥 Créditos")

        layout.addWidget(tabs)
        self.setLayout(layout)

    # ==========================================================
    # MÉTODOS PARA RELATÓRIOS ESPECÍFICOS
    # ==========================================================

    def criar_aba_relatorios_especificos(self):
        """Cria a aba com os relatórios específicos (barramento e parcelamento) - APENAS DADOS"""
        widget = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)

        # ==========================================================
        # RELATÓRIO DE BARRAMENTO - CAMPOS ORGANIZADOS
        # ==========================================================
        grupo_barramento = QGroupBox("💧 DADOS PARA RELATÓRIO DE BARRAMENTO")
        grupo_barramento.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 2px solid #006b3f;
                border-radius: 8px;
                margin-top: 15px;
            }
            QGroupBox::title {
                color: #006b3f;
                font-weight: bold;
                font-size: 12pt;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
            }
        """)

        barramento_layout = QVBoxLayout()
        barramento_layout.setSpacing(10)

        # -------- 1. DADOS GERAIS DA OCORRÊNCIA --------
        sub_titulo1 = QLabel("📋 DADOS GERAIS DA OCORRÊNCIA")
        sub_titulo1.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
        """)
        barramento_layout.addWidget(sub_titulo1)

        grid1 = QGridLayout()
        grid1.setVerticalSpacing(8)
        grid1.setHorizontalSpacing(15)

        # Data da Ocorrência
        lbl1 = QLabel("📅 Data da Ocorrência:")
        lbl1.setStyleSheet("font-weight: 500;")
        campo1 = QLineEdit()
        campo1.setPlaceholderText("Digite a data da ocorrência")
        campo1.setText(datetime.now().strftime("%d/%m/%Y"))
        self.campos_barramento["data_ocorrencia"] = campo1
        grid1.addWidget(lbl1, 0, 0)
        grid1.addWidget(campo1, 0, 1)

        # Manifestação
        lbl2 = QLabel("📝 Manifestação:")
        lbl2.setStyleSheet("font-weight: 500;")
        campo2 = QLineEdit()
        campo2.setPlaceholderText("Digite a manifestação")
        self.campos_barramento["manifestacao"] = campo2
        grid1.addWidget(lbl2, 1, 0)
        grid1.addWidget(campo2, 1, 1)

        barramento_layout.addLayout(grid1)

        # -------- 2. DESCRIÇÃO --------
        sub_titulo2 = QLabel("📄 DESCRIÇÃO")
        sub_titulo2.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
            margin-top: 10px;
        """)
        barramento_layout.addWidget(sub_titulo2)

        grid2 = QGridLayout()
        grid2.setVerticalSpacing(8)
        grid2.setHorizontalSpacing(15)

        # Fato Denunciado
        lbl3 = QLabel("📄 Fato Denunciado:")
        lbl3.setStyleSheet("font-weight: 500;")
        campo3 = QTextEdit()
        campo3.setPlaceholderText("Digite o fato denunciado")
        campo3.setMaximumHeight(60)
        self.campos_barramento["fato_denunciado"] = campo3
        grid2.addWidget(lbl3, 0, 0)
        grid2.addWidget(campo3, 0, 1)

        # Texto INÃ PR
        lbl4 = QLabel("📝 Texto INÃ PR:")
        lbl4.setStyleSheet("font-weight: 500;")
        campo4 = QTextEdit()
        campo4.setPlaceholderText("Digite o texto INÃ PR")
        campo4.setMaximumHeight(60)
        self.campos_barramento["texto_ina_pr"] = campo4
        grid2.addWidget(lbl4, 1, 0)
        grid2.addWidget(campo4, 1, 1)

        barramento_layout.addLayout(grid2)

        # -------- 3. GEOPROCESSAMENTO --------
        sub_titulo3 = QLabel("🗺️ GEOPROCESSAMENTO")
        sub_titulo3.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
            margin-top: 10px;
        """)
        barramento_layout.addWidget(sub_titulo3)

        grid3 = QGridLayout()
        grid3.setVerticalSpacing(8)
        grid3.setHorizontalSpacing(15)

        # Data do Sobrevoo
        lbl5 = QLabel("📅 Data do Sobrevoo:")
        lbl5.setStyleSheet("font-weight: 500;")
        campo5 = QLineEdit()
        campo5.setPlaceholderText("Digite a data do sobrevoo")
        self.campos_barramento["data_sobrevoo"] = campo5
        grid3.addWidget(lbl5, 0, 0)
        grid3.addWidget(campo5, 0, 1)

        # Mapa Geral
        lbl6 = QLabel("🗺️ Mapa Geral:")
        lbl6.setStyleSheet("font-weight: 500;")
        campo6 = QLineEdit()
        campo6.setPlaceholderText("Digite o nome do mapa geral")
        self.campos_barramento["mapa_geral"] = campo6
        grid3.addWidget(lbl6, 1, 0)
        grid3.addWidget(campo6, 1, 1)

        # Mapa Temporal
        lbl7 = QLabel("🗺️ Mapa Temporal:")
        lbl7.setStyleSheet("font-weight: 500;")
        campo7 = QLineEdit()
        campo7.setPlaceholderText("Digite o nome do mapa temporal")
        self.campos_barramento["mapa_temporal"] = campo7
        grid3.addWidget(lbl7, 2, 0)
        grid3.addWidget(campo7, 2, 1)

        barramento_layout.addLayout(grid3)

        # -------- 4. ÁREAS --------
        sub_titulo4 = QLabel("🌿 ÁREAS")
        sub_titulo4.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
            margin-top: 10px;
        """)
        barramento_layout.addWidget(sub_titulo4)

        grid4 = QGridLayout()
        grid4.setVerticalSpacing(8)
        grid4.setHorizontalSpacing(15)

        # Área do Barramento
        lbl8 = QLabel("🌊 Área do Barramento (ha):")
        lbl8.setStyleSheet("font-weight: 500;")
        campo8 = QLineEdit()
        campo8.setPlaceholderText("Digite a área do barramento em hectares")
        area_app = self.areas_inputs.get("app", {}).get("area", QLineEdit()).text()
        if area_app:
            campo8.setText(area_app)
        self.campos_barramento["area_barramento"] = campo8
        grid4.addWidget(lbl8, 0, 0)
        grid4.addWidget(campo8, 0, 1)

        # Área de Supressão
        lbl9 = QLabel("🌿 Área de Supressão (ha):")
        lbl9.setStyleSheet("font-weight: 500;")
        campo9 = QLineEdit()
        campo9.setPlaceholderText("Digite a área de supressão em hectares")
        area_fora = self.areas_inputs.get("fora", {}).get("area", QLineEdit()).text()
        if area_fora:
            campo9.setText(area_fora)
        self.campos_barramento["area_supressao"] = campo9
        grid4.addWidget(lbl9, 1, 0)
        grid4.addWidget(campo9, 1, 1)

        barramento_layout.addLayout(grid4)

        # -------- 5. VALORES --------
        sub_titulo5 = QLabel("💰 VALORES")
        sub_titulo5.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
            margin-top: 10px;
        """)
        barramento_layout.addWidget(sub_titulo5)

        grid5 = QGridLayout()
        grid5.setVerticalSpacing(8)
        grid5.setHorizontalSpacing(15)

        # Valor Recursos Hídricos
        lbl10 = QLabel("💰 Valor Recursos Hídricos:")
        lbl10.setStyleSheet("font-weight: 500;")
        campo10 = QLineEdit()
        campo10.setPlaceholderText("Digite o valor dos recursos hídricos")
        self.campos_barramento["valor_recursos_hidricos"] = campo10
        grid5.addWidget(lbl10, 0, 0)
        grid5.addWidget(campo10, 0, 1)

        barramento_layout.addLayout(grid5)

        # ==========================================================
        # 📌 CAMPOS QUE SERÃO PREENCHIDOS AUTOMATICAMENTE
        # ==========================================================
        info_frame = QFrame()
        info_frame.setStyleSheet("""
            QFrame {
                background-color: #e8f5e9;
                border-radius: 8px;
                padding: 12px;
                margin-top: 10px;
            }
        """)
        info_layout = QVBoxLayout()

        info_titulo = QLabel("🔗 DADOS COLETADOS AUTOMATICAMENTE")
        info_titulo.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt;")
        info_layout.addWidget(info_titulo)

        info_detalhes = QLabel("""
        ✅ Auto Art. 66 → do Art. 66
        ✅ Embargo Art. 66 → do Art. 66
        ✅ Valor AI 66 → do Art. 66

        ✅ Auto Art. 43 → do Art. 43
        ✅ Embargo Art. 43 → do Art. 43
        ✅ Valor AI 43 → do Art. 43

        ✅ Intervalo da Supressão → dos Dados Gerais
        ✅ Área de Supressão Arredondada → calculada automaticamente
        """)
        info_detalhes.setStyleSheet("color: #2c3e50; font-size: 10pt; line-height: 1.8;")
        info_layout.addWidget(info_detalhes)

        info_frame.setLayout(info_layout)
        barramento_layout.addWidget(info_frame)

        grupo_barramento.setLayout(barramento_layout)
        layout.addWidget(grupo_barramento)

        # ==========================================================
        # RELATÓRIO DE PARCELAMENTO - CAMPOS ORGANIZADOS
        # ==========================================================
        grupo_parcelamento = QGroupBox("🏘️ DADOS PARA RELATÓRIO DE PARCELAMENTO")
        grupo_parcelamento.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 2px solid #006b3f;
                border-radius: 8px;
                margin-top: 15px;
            }
            QGroupBox::title {
                color: #006b3f;
                font-weight: bold;
                font-size: 12pt;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
            }
        """)

        parcelamento_layout = QVBoxLayout()
        parcelamento_layout.setSpacing(10)

        # ==========================================================
        # 📌 CAMPOS ORGANIZADOS
        # ==========================================================

        # -------- 1. DADOS GERAIS DA OCORRÊNCIA --------
        sub_titulo_parc = QLabel("📋 DADOS GERAIS DA OCORRÊNCIA")
        sub_titulo_parc.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
        """)
        parcelamento_layout.addWidget(sub_titulo_parc)

        grid_parc = QGridLayout()
        grid_parc.setVerticalSpacing(8)
        grid_parc.setHorizontalSpacing(15)

        # Data da Ocorrência
        lbl_parc1 = QLabel("📅 Data da Ocorrência:")
        lbl_parc1.setStyleSheet("font-weight: 500;")
        campo_parc1 = QLineEdit()
        campo_parc1.setPlaceholderText("Digite a data da ocorrência")
        campo_parc1.setText(datetime.now().strftime("%d/%m/%Y"))
        self.campos_parcelamento["parc_data_ocorrencia"] = campo_parc1
        grid_parc.addWidget(lbl_parc1, 0, 0)
        grid_parc.addWidget(campo_parc1, 0, 1)

        # Manifestação
        lbl_parc2 = QLabel("📝 Manifestação:")
        lbl_parc2.setStyleSheet("font-weight: 500;")
        campo_parc2 = QLineEdit()
        campo_parc2.setPlaceholderText("Digite a manifestação")
        self.campos_parcelamento["parc_manifestacao"] = campo_parc2
        grid_parc.addWidget(lbl_parc2, 1, 0)
        grid_parc.addWidget(campo_parc2, 1, 1)

        parcelamento_layout.addLayout(grid_parc)

        # -------- 2. DESCRIÇÃO --------
        sub_titulo_parc2 = QLabel("📄 DESCRIÇÃO")
        sub_titulo_parc2.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
            margin-top: 10px;
        """)
        parcelamento_layout.addWidget(sub_titulo_parc2)

        grid_parc2 = QGridLayout()
        grid_parc2.setVerticalSpacing(8)
        grid_parc2.setHorizontalSpacing(15)

        # Fato Denunciado
        lbl_parc3 = QLabel("📄 Fato Denunciado:")
        lbl_parc3.setStyleSheet("font-weight: 500;")
        campo_parc3 = QTextEdit()
        campo_parc3.setPlaceholderText("Digite o fato denunciado")
        campo_parc3.setMaximumHeight(60)
        self.campos_parcelamento["parc_fato_denunciado"] = campo_parc3
        grid_parc2.addWidget(lbl_parc3, 0, 0)
        grid_parc2.addWidget(campo_parc3, 0, 1)

        parcelamento_layout.addLayout(grid_parc2)

        # -------- 3. DADOS ESPECÍFICOS --------
        sub_titulo_parc3 = QLabel("📊 DADOS ESPECÍFICOS")
        sub_titulo_parc3.setStyleSheet("""
            font-weight: bold; 
            color: #006b3f; 
            font-size: 11pt; 
            padding: 5px 0;
            border-bottom: 1px solid #e1e4e8;
            margin-top: 10px;
        """)
        parcelamento_layout.addWidget(sub_titulo_parc3)

        grid_parc3 = QGridLayout()
        grid_parc3.setVerticalSpacing(8)
        grid_parc3.setHorizontalSpacing(15)

        # Quantidade de Lotes
        lbl_parc4 = QLabel("📊 Quantidade de Lotes:")
        lbl_parc4.setStyleSheet("font-weight: 500;")
        campo_parc4 = QLineEdit()
        campo_parc4.setPlaceholderText("Digite a quantidade de lotes")
        self.campos_parcelamento["parc_qtd_lotes"] = campo_parc4
        grid_parc3.addWidget(lbl_parc4, 0, 0)
        grid_parc3.addWidget(campo_parc4, 0, 1)

        # Informação da APP
        lbl_parc5 = QLabel("🌳 Informação da APP:")
        lbl_parc5.setStyleSheet("font-weight: 500;")
        campo_parc5 = QLineEdit()
        campo_parc5.setPlaceholderText("Digite a informação da APP (ex: APP do Rio dos Bois)")
        self.campos_parcelamento["parc_app_info"] = campo_parc5
        grid_parc3.addWidget(lbl_parc5, 1, 0)
        grid_parc3.addWidget(campo_parc5, 1, 1)

        parcelamento_layout.addLayout(grid_parc3)

        # ==========================================================
        # 📌 CAMPOS QUE SERÃO PREENCHIDOS AUTOMATICAMENTE DO ART. 66
        # ==========================================================
        info_frame_parc = QFrame()
        info_frame_parc.setStyleSheet("""
            QFrame {
                background-color: #e8f5e9;
                border-radius: 8px;
                padding: 12px;
                margin-top: 10px;
            }
        """)
        info_layout_parc = QVBoxLayout()

        info_titulo_parc = QLabel("🔗 DADOS COLETADOS AUTOMATICAMENTE DO ART. 66")
        info_titulo_parc.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt;")
        info_layout_parc.addWidget(info_titulo_parc)

        info_detalhes_parc = QLabel("""
        ✅ Auto de Infração → do Art. 66
        ✅ Embargo → do Art. 66
        ✅ Tipo de Infrator → do Art. 66
        ✅ Faixa de Receita → do Art. 66
        ✅ Nível de Gravidade → do Art. 66
        ✅ Percentual Aplicado → do Art. 66
        ✅ Valor da Multa → do Art. 66
        """)
        info_detalhes_parc.setStyleSheet("color: #2c3e50; font-size: 10pt; line-height: 1.6;")
        info_layout_parc.addWidget(info_detalhes_parc)

        info_frame_parc.setLayout(info_layout_parc)
        parcelamento_layout.addWidget(info_frame_parc)

        grupo_parcelamento.setLayout(parcelamento_layout)
        layout.addWidget(grupo_parcelamento)

        # ==========================================================
        # 📌 MENSAGEM INFORMATIVA
        # ==========================================================
        msg_frame = QFrame()
        msg_frame.setStyleSheet("""
            QFrame {
                background-color: #fff3cd;
                border: 1px solid #ffeeba;
                border-radius: 8px;
                padding: 12px;
                margin-top: 10px;
            }
        """)
        msg_layout = QVBoxLayout()

        msg_label = QLabel("💡 Os relatórios são gerados na aba 'Relatórios'")
        msg_label.setStyleSheet("""
            font-weight: bold;
            color: #856404;
            font-size: 11pt;
        """)
        msg_label.setAlignment(Qt.AlignCenter)

        msg_sub = QLabel("""
        Preencha os dados acima e depois vá na aba 'Relatórios', 
        selecione o relatório desejado e clique em 'GERAR RELATÓRIOS SELECIONADOS'
        """)
        msg_sub.setStyleSheet("color: #856404; font-size: 10pt;")
        msg_sub.setAlignment(Qt.AlignCenter)
        msg_sub.setWordWrap(True)

        msg_layout.addWidget(msg_label)
        msg_layout.addWidget(msg_sub)
        msg_frame.setLayout(msg_layout)
        layout.addWidget(msg_frame)

        layout.addStretch()

        container.setLayout(layout)
        scroll.setWidget(container)

        final_layout = QVBoxLayout()
        final_layout.addWidget(scroll)
        widget.setLayout(final_layout)

        return widget
    def preencher_campos_barramento(self):
        """Preenche automaticamente os campos do barramento com dados existentes"""
        # Pega dados das outras abas
        dados = self.gerenciador.obter_dados_comuns()

        # Preenche campos
        if 'manifestacao' in self.campos_barramento:
            self.campos_barramento['manifestacao'].setText(
                f"Manifestação {dados.get('processo', 'NÃO INFORMADO')}"
            )

        if 'fato_denunciado' in self.campos_barramento:
            self.campos_barramento['fato_denunciado'].setText(
                dados.get('observacoes', 'NÃO INFORMADO')
            )

        # Pega áreas das abas
        area_app = self.areas_inputs.get("app", {}).get("area", QLineEdit()).text()
        area_fora = self.areas_inputs.get("fora", {}).get("area", QLineEdit()).text()

        if area_app and 'area_barramento' in self.campos_barramento:
            self.campos_barramento['area_barramento'].setText(area_app)

        if area_fora and 'area_supressao' in self.campos_barramento:
            self.campos_barramento['area_supressao'].setText(area_fora)
            # Calcula arredondamento
            try:
                area_float = float(area_fora.replace(',', '.'))
                area_arredondada = math.ceil(area_float)
                if 'area_supressao_arredondada' in self.campos_barramento:
                    self.campos_barramento['area_supressao_arredondada'].setText(str(area_arredondada))
            except:
                pass

        # Auto e Embargo - tenta pegar das áreas
        if 'auto_barramento' in self.campos_barramento:
            auto_app = self.areas_inputs.get("app", {}).get("auto", QTextEdit()).toPlainText()
            if auto_app:
                self.campos_barramento['auto_barramento'].setText(auto_app)

        if 'embargo_barramento' in self.campos_barramento:
            embargo_app = self.areas_inputs.get("app", {}).get("embargo", QTextEdit()).toPlainText()
            if embargo_app:
                self.campos_barramento['embargo_barramento'].setText(embargo_app)

        # Valores das multas
        if 'valor_ai_66' in self.campos_barramento:
            self.campos_barramento['valor_ai_66'].setText("3000,00")

        if 'valor_ai_43' in self.campos_barramento:
            self.campos_barramento['valor_ai_43'].setText("5012,50")

        if 'valor_recursos_hidricos' in self.campos_barramento:
            self.campos_barramento['valor_recursos_hidricos'].setText("902,25")

        QMessageBox.information(self, "Sucesso", "Campos preenchidos automaticamente com os dados disponíveis!")


    def gerar_relatorio_barramento_ui(self):
        """Gera o relatório de barramento a partir da interface"""
        try:
            # Coleta dados dos campos
            dados = {}
            for nome, campo in self.campos_barramento.items():
                if hasattr(campo, 'toPlainText'):
                    dados[nome] = campo.toPlainText()
                elif hasattr(campo, 'currentText'):
                    dados[nome] = campo.currentText()
                else:
                    dados[nome] = campo.text()

            # Valida campos obrigatórios
            obrigatorios = ['data_ocorrencia', 'manifestacao', 'auto_barramento']
            for campo in obrigatorios:
                if not dados.get(campo):
                    QMessageBox.warning(self, "Aviso", f"O campo '{campo.replace('_', ' ').title()}' é obrigatório!")
                    return

            # Gera o relatório usando o gerenciador
            relatorio = self.gerenciador.gerar_relatorio_barramento(dados)

            # Salva o arquivo
            pasta = QFileDialog.getExistingDirectory(self, "Escolha a pasta para salvar")
            if not pasta:
                return

            nome_arquivo = f"Relatorio_Barramento_{dados.get('imovel', 'NAO_INFORMADO')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            caminho = os.path.join(pasta, nome_arquivo)

            with open(caminho, 'w', encoding='utf-8') as f:
                f.write(relatorio)

            QMessageBox.information(self, "Sucesso", f"Relatório gerado com sucesso!\n{caminho}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar relatório: {str(e)}")

    def gerar_relatorio_parcelamento(self, dados_especificos):
        """Gera o relatório de parcelamento completo usando dados das abas"""
        dados = self.gerenciador.obter_dados_comuns()
        dados.update(dados_especificos)

        print("=" * 60)
        print("DADOS RECEBIDOS EM gerar_relatorio_parcelamento:")
        print(f"parc_fato_denunciado: {dados.get('parc_fato_denunciado', 'NÃO ENCONTRADO')[:200]}...")
        print(f"parc_manifestacao: {dados.get('parc_manifestacao', 'NÃO ENCONTRADO')}")
        print(f"parc_data_ocorrencia: {dados.get('parc_data_ocorrencia', 'NÃO ENCONTRADO')}")
        print(f"parc_qtd_lotes: {dados.get('parc_qtd_lotes', 'NÃO ENCONTRADO')}")
        print("=" * 60)
        # ==========================================================
        # BUSCA DADOS DO ART. 66 AUTOMATICAMENTE
        # ==========================================================
        import re

        # 1. Buscar Auto do Art. 66
        auto_parcelamento = ""
        if "art66" in self.artigos_inputs:
            auto_widget = self.artigos_inputs["art66"].get("auto")
            if auto_widget and hasattr(auto_widget, 'text'):
                auto_parcelamento = auto_widget.text().strip()

        if not auto_parcelamento:
            auto_parcelamento = "NÃO INFORMADO"

        # 2. Buscar Embargo do Art. 66
        embargo_parcelamento = ""
        if "art66" in self.artigos_inputs:
            embargo_widget = self.artigos_inputs["art66"].get("embargo")
            if embargo_widget and hasattr(embargo_widget, 'text'):
                embargo_parcelamento = embargo_widget.text().strip()

        if not embargo_parcelamento:
            embargo_parcelamento = "NÃO INFORMADO"

        # 3. Buscar Tipo de Infrator do Art. 66
        tipo_infrator = "Pessoa Física"
        if "art66" in self.artigos_inputs:
            tipo_widget = self.artigos_inputs["art66"].get("tipo_infrator")
            if tipo_widget and hasattr(tipo_widget, 'currentText'):
                tipo_infrator = tipo_widget.currentText()

        if not tipo_infrator or tipo_infrator == "Selecione...":
            tipo_infrator = "Pessoa Física"

        # 4. Buscar Faixa de Receita do Art. 66
        faixa_receita = "Faixa A"
        if "art66" in self.artigos_inputs:
            faixa_widget = self.artigos_inputs["art66"].get("faixa_receita")
            if faixa_widget and hasattr(faixa_widget, 'currentText'):
                faixa_receita = faixa_widget.currentText()
                if " - " in faixa_receita:
                    faixa_receita = faixa_receita.split(" - ")[0]

        # 5. Buscar Nível de Gravidade do Art. 66
        nivel_gravidade = "Nível C"
        if "art66" in self.artigos_inputs:
            pontuacao_widget = self.artigos_inputs["art66"].get("pontuacao")
            if pontuacao_widget and hasattr(pontuacao_widget, 'text'):
                pontuacao_text = pontuacao_widget.text()
                if "Nível" in pontuacao_text:
                    nivel_match = re.search(r'(Nível\s+[A-E])', pontuacao_text)
                    if nivel_match:
                        nivel_gravidade = nivel_match.group(1)

        # 6. Buscar Percentual do Art. 66
        percentual = "0,03%"
        if "art66" in self.artigos_inputs:
            percentual_widget = self.artigos_inputs["art66"].get("percentual")
            if percentual_widget and hasattr(percentual_widget, 'text'):
                percentual_text = percentual_widget.text()
                perc_match = re.search(r'(\d+[.,]\d+%?)', percentual_text)
                if perc_match:
                    percentual = perc_match.group(1)

        # 7. Buscar Valor da Multa do Art. 66
        valor_multa = 3500.00
        if "art66" in self.artigos_inputs:
            valor_widget = self.artigos_inputs["art66"].get("valor")
            if valor_widget and hasattr(valor_widget, 'text'):
                valor_text = valor_widget.text().strip()
                if valor_text:
                    numeros = re.findall(r'[\d\.,]+', valor_text)
                    if numeros:
                        try:
                            valor_str = numeros[0].replace('.', '').replace(',', '.')
                            valor_multa = float(valor_str)
                        except:
                            pass

        # ==========================================================
        # DADOS DA ABA "DADOS PARA O RELATÓRIO"
        # ==========================================================
        # Data da Ocorrência
        data_ocorrencia = dados.get('parc_data_ocorrencia', datetime.now().strftime("%d/%m/%Y"))

        # Manifestação
        manifestacao = dados.get('parc_manifestacao', 'NÃO INFORMADO')

        
        # Fato Denunciado
        fato_denunciado = dados.get('parc_fato_denunciado', 'NÃO INFORMADO')

        # ==========================================================
        # DADOS GERAIS (prioriza interface)
        # ==========================================================
        # PROPRIETÁRIO
        proprietario = dados.get('proprietario', '').strip()
        if not proprietario or proprietario == 'NÃO INFORMADO':
            proprietario = "NÃO INFORMADO"

        # CPF
        cpf = dados.get('cpf', '').strip()
        if not cpf or cpf == 'NÃO INFORMADO':
            cpf = "NÃO INFORMADO"

        # MUNICÍPIO
        municipio = dados.get('municipio', '').strip()
        if not municipio or municipio == 'NÃO INFORMADO':
            municipio = "NÃO INFORMADO"

        # IMÓVEL
        imovel = dados.get('imovel', '').strip()
        if not imovel or imovel == 'NÃO INFORMADO':
            imovel = "NÃO INFORMADO"

        # COORDENADAS
        coordenadas = dados.get('coordenadas', '').strip()
        if not coordenadas or coordenadas == 'NÃO INFORMADO':
            coordenadas = "NÃO INFORMADO"

        # PROCESSO, OS, CAR
        processo = dados.get('processo', 'NÃO INFORMADO').strip() or "NÃO INFORMADO"
        os_num = dados.get('os', 'NÃO INFORMADO').strip() or "NÃO INFORMADO"
        car = dados.get('car', 'NÃO INFORMADO').strip() or "NÃO INFORMADO"

        # ==========================================================
        # DADOS DO USUÁRIO (criar campos na aba Dados para o relatório)
        # ==========================================================
        # Quantidade de Lotes - usar valor da interface ou extrair do texto
        qtd_lotes = dados.get('parc_qtd_lotes', 'NÃO INFORMADO')
        if not qtd_lotes or qtd_lotes == 'NÃO INFORMADO':
            # Tenta extrair do texto do fato denunciado
            lotes_match = re.search(r'(\d+)\s*lotes', fato_denunciado, re.IGNORECASE)
            if lotes_match:
                qtd_lotes = lotes_match.group(1)

        # Informações sobre casas - extrair do texto
        casas_info = "NÃO INFORMADO"
        casas_match = re.search(r'(\d+)\s*casas\s+já\s+construídas', fato_denunciado, re.IGNORECASE)
        if casas_match:
            casas_info = casas_match.group(0)

        # APP - usar valor da interface ou extrair do texto
        app_info = dados.get('parc_app_info', 'NÃO INFORMADO')
        if not app_info or app_info == 'NÃO INFORMADO':
            if "Rio dos Bois" in fato_denunciado:
                app_info = "APP do Rio dos Bois"
            elif "APP" in fato_denunciado:
                app_info = "APP"

        # Fração mínima - extrair do texto
        fracao_minima = "NÃO INFORMADO"
        fracao_match = re.search(r'área mínima de\s+(\d+\s*\([^\)]+\)\s*hectares)', fato_denunciado, re.IGNORECASE)
        if fracao_match:
            fracao_minima = fracao_match.group(1)

        # Microrregião - extrair do texto
        microrregiao = "NÃO INFORMADO"
        micro_match = re.search(r'região\s*(\d+\s*\([^\)]+\))', fato_denunciado, re.IGNORECASE)
        if micro_match:
            microrregiao = micro_match.group(1)

        # ==========================================================
        # FORMATAÇÃO DA DATA
        # ==========================================================
        data_fim = datetime.now().strftime("%d/%m/%Y")

        # ==========================================================
        # CONSTRUÇÃO DO RELATÓRIO
        # ==========================================================
        relatorio = f"""RELATÓRIO DE FISCALIZAÇÃO - PARCELAMENTO IMOBILIÁRIO IRREGULAR

    Motivo/Objetivo da Fiscalização
    Trata-se de denúncia/comunicação recepcionada junto ao Sistema de Ouvidoria do Estado de Goiás da Controladoria Geral do Estado, através da qual o {fato_denunciado}

    A Ouvidoria Setorial após análise prévia da manifestação elaborou o Despacho nº {processo}

    Em cumprimento a OS SEMAD/GO de nº {os_num}, no dia {data_ocorrencia}, esta equipe de fiscalização deslocou-se até o município de {municipio} para averiguar as informações prestadas na Manifestação {manifestacao}. Cuja a descrição: {fato_denunciado}

    Para cumprir com os objetivos estabelecidos, foram empregadas diversas metodologias de fiscalização, incluindo análise de imagens de satélite, observação in loco e documental, visando uma apuração detalhada e precisa da situação.

    Tipo da Ação:
    ( ) Análise/Fiscalização Processual
    ( ) Fiscalização (remota) – §4° do Art.36 da Lei Estadual 20.694/2019
    ( ) Fiscalização in loco
    ( ) Fiscalização em atividade sem licença.

    Motivação:
    ( ) Acompanhamento
    ( ) Análise Complementar (juntada de documento após fiscalização)
    ( ) Análise de Atendimento de Notificação
    (X) Denúncia/Ouvidoria/INÃ – {manifestacao}
    ( ) Análise de Processos solicitados por órgãos externos – Processo:
    ( ) LAI - Lei de Acesso à Informação – Processo:

    Localização da Atividade:
    ( ) Zona Urbana: {imovel}, {municipio}, {coordenadas}.
    (X) Zona Rural: {imovel}, Coordenadas: {coordenadas}

    Descrição da atividade fiscalizatória
    1. HISTÓRICO PROCESSUAL:
    SGA: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    IPÊ: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    SEI: ( ) Nada consta ( ) Não se aplica (X) Processo(s)/Cadastro(s): {dados.get('processo_sei', 'NÃO INFORMADO')}
    INÃ: ( ) Nada consta ( ) Não se aplica (X) Processo(s)/Cadastro(s): {dados.get('processo_ina', 'NÃO INFORMADO')}
    WebOutorga / Veredas: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    SIGA: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    SICAR: ( ) Nada consta ( ) Não se aplica (X) Processo(s)/Cadastro(s): {car}
    SEISB: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):

    2. CONSTATAÇÕES
    2.1 Constatações Técnicas da Fiscalização
    () A equipe foi recebida por: {proprietario}, CPF: {cpf};
    ( ) Ninguém foi encontrado no local.

    Observações detalhadas da fiscalização:

    Abertura de ruas de acesso (estrada vicinal);
    Sistema de rede de energia com 01 transformador que faz a distribuição nos lotes;
    Uso água é por meio de poço perfurado individualmente;
    Divisão dos lotes (cercas e alambrados);
    {qtd_lotes} lotes ao todo;
    As divisões foram efetuadas de tal forma para que todos tenham acesso ao Rio dos Bois;
    {casas_info};

    No local, a maioria das casas estava com a porteira de acesso trancada. Em uma das propriedades, a equipe foi recebida por {proprietario}, que adquiriu um lote de aproximadamente 1500m². A moradora estava portando o contrato de compra e venda celebrado entre as partes com assinatura registrada no Tabelionato de Notas, de Protesto de Títulos, Tabelionato e Oficialato de Registro de Contratos Marítimos, de Registro de Imóveis, de Registro de Títulos e Documentos, Civil das Pessoas Jurídicas, Civil das Pessoas Naturais e de Interdições e Tutelas, de {municipio} - Goiás.

    Conforme a análise dos dados através de imagens, o início do parcelamento é anterior a agosto de 2019, considerando que, ao longo dos anos, houve um aumento progressivo na instalação e construção de moradias no condomínio.

    Tendo em vista que o empreendimento foi instalado sem a devida autorização ou licença ambiental, foram lavrados em nome de {proprietario}, CPF: {cpf}, o Auto de Infração n° {auto_parcelamento}, Termo de Embargo n° {embargo_parcelamento}, acompanhados da respectiva Carta Convite para Autocomposição. Todos os autos lavrados e encaminhados via Carta registrada, conforme dados coletados no sistema INFOSEG, com aviso de recebimento.

    1. Que atividade(s) é(são) realizada(s) no local?
    No local é realizada a atividade de parcelamento/loteamento.

    2. A(s) atividade(s) é(são) utilizadora(s) de recursos ambientais, consideradas efetiva ou potencialmente poluidoras, ou capazes de, sob qualquer forma, causar degradação ambiental?
    Sim. As atividades de parcelamento (uso do solo, em imóveis rurais, decorrente de desmembramento imobiliário), são consideradas potencialmente poluidoras do meio ambiente sendo, portanto, passíveis de licenciamento ambiental.

    3. A(s) atividade(s) é(são) licenciada(s)/autorizada(s)?
    Não. A atividade está sendo realizada sem a devida autorização ou licenciamento ambiental.

    4. Há quanto tempo a(s) atividade(s) é(são) realizada(s) no local?
    Atividade iniciada antes de agosto de 2019.

    5. Quem é(são) o(s) responsável(is) pela(s) atividade(s)?
    {proprietario}, CPF: {cpf}

    6. Qual o endereço e coordenadas geográficas do local onde é(são) realizada(s) a(s) atividade(s)?
    {imovel}, no município de {municipio}, no entorno da Coordenada Geográfica SIRGAS 2000: {coordenadas}

    7. A(s) atividade(s) é(são) desenvolvida(s) em áreas protegidas (APP, Reserva Legal e/ou Unidade de Conservação)?
    Sim, {app_info}.

    8. A(s) atividade(s) causou(ram) ou está(ão) causando degradação ambiental?
    Sim. A subdivisão de propriedades rurais para a formação de chácaras e condomínios tem causado sérios problemas ambientais. Esse processo resulta na destruição de habitats naturais, afetando a biodiversidade e levando à extinção de espécies. A fragmentação das terras também compromete os recursos hídricos, pois a impermeabilização do solo aumenta o escoamento superficial, causando erosão e poluição de rios e lagos. A poluição gerada pelo desmembramento de propriedades inclui a contaminação dos corpos d'água por esgoto doméstico e produtos químicos utilizados em jardins e áreas de lazer. A qualidade do ar também pode ser comprometida pelo aumento de emissões de veículos e pela queima de resíduos. O solo, sem a cobertura vegetal adequada, torna-se mais suscetível à erosão, perdendo nutrientes e capacidade de retenção de água. Isso contribui para a degradação do solo e pode levar à desertificação. As mudanças no uso do solo também afetam o microclima local, alterando a temperatura e a umidade do ar.

    9. Que medidas foram ou devem ser adotadas para cessar a degradação?
    Foram adotadas as seguintes medidas administrativas:
    Lavratura do Auto de Infração n.º {auto_parcelamento};
    Emissão do Termo de Embargo n.º {embargo_parcelamento}, determinando a suspensão imediata da atividade irregular.

    10. A degradação ambiental comporta recuperação?
    Sim. A degradação ambiental identificada é passível de recuperação. A viabilidade e o tempo necessários para a recomposição dos recursos hídricos e da funcionalidade do ecossistema afetado dependerão da extensão do impacto, das características hidrogeológicas da área e das medidas de mitigação implementadas.

    11. Que medidas foram ou devem ser adotadas para promover a recuperação ambiental?
    As ações de recuperação ambiental deverão ser definidas no âmbito do processo de regularização ambiental do empreendimento.

    Da valoração

    O artigo 66 do Decreto nº 6.514/2008 estabelece a valoração de Multa de R$ 500,00 (quinhentos reais) a R$ 10.000.000,00 (dez milhões de reais):
    Art. 66. Construir, reformar, ampliar, instalar ou fazer funcionar estabelecimentos, atividades, obras ou serviços utilizadores de recursos ambientais, considerados efetiva ou potencialmente poluidores, sem licença ou autorização dos órgãos ambientais competentes, em desacordo com a licença obtida ou contrariando as normas legais e regulamentos pertinentes:
    Multa de R$ 500,00 (quinhentos reais) a R$ 10.000.000,00 (dez milhões de reais).

    Para fixação do valor referente a infração foram utilizados os critérios definidos na ORIENTAÇÃO NORMATIVA SEMAD Nº 01/2024 (Define parâmetros para a fixação das multas abertas, para a aplicação de sanções e medidas administrativas cautelares no âmbito da apuração de infrações ambientais, bem como critérios para o agravamento e a atenuação das sanções administrativas decorrentes de infrações ambientais), considerando o nível de gravidade da infração e identificação da capacidade econômica da seguinte maneira:
    - Motivo da Infração: Obtenção de vantagem pecuniária (15)
    - Consequência para o meio ambiente: Moderada (30)
    - Consequência para a saúde pública ou para a socioeconomia da área de abrangência do fato: Não Houve (0)
    Somatório dos valores desta etapa: (45) - {nivel_gravidade}

    Situação econômica: Não foi possível aferir a situação econômica do infrator. Por conta disso, foi utilizado Situação econômica - Receita Mensal - {tipo_infrator} - Receita bruta: mensal de até 1 salário mínimo ({faixa_receita}) = {nivel_gravidade}: Mínimo + {percentual} até 1,0% do teto.

    VALORAÇÃO: R$ 500,00 + ({percentual} X R$ 10.000.000,00) = R$ {self.gerenciador._formatar_moeda_br(valor_multa)} ({self.gerenciador._numero_extenso(valor_multa)})

    Considerando-se que foi utilizado a classe de menor valor para a situação econômica do infrator, que pode não ser condizente com a realidade, sugere-se na Autocomposição/julgamento, a aferição e reenquadramento se for o caso conforme Art. 8º, §2º ON 01/2024 SEMAD.
    Art. 8º Em se tratando de pessoa física, a situação econômica do infrator será determinada pelos critérios estabelecidos no Quadro 3 do Anexo único, mediante a classificação em faixas definidas conforme receita bruta anual do infrator, assim estabelecidas:
    I - receita bruta mensal de até 1 salário mínimo;
    II - receita bruta mensal, superior ao limite anterior até 3 salários mínimos;
    III - receita bruta mensal, superior ao limite anterior até 10 salários mínimos;
    IV - receita bruta mensal, superior ao limite anterior até 30 salários mínimos;
    V - receita bruta mensal, superior ao limite anterior até 45 salários mínimos; e
    VI - receita bruta mensal, superior ao limite anterior.
    § 1º Em se tratando de pessoa física serão considerados os rendimentos indicados em qualquer documento válido para comprovação de renda.
    § 2º A autoridade julgadora competente bem como os facilitadores em sede de audiências de autocomposição deverão rever o enquadramento do infrator quanto a sua situação econômica, caso conste no relatório de fiscalização que esta não tenha sido possível aferir.

    INFORMAÇÕES COMPLEMENTARES:

    LEI COMPLEMENTAR Nº 140, DE 8 DE DEZEMBRO DE 2011
    No artigo 3º da Resolução CEMAm 259/2024, especifica quais os parâmetros e requisitos o município deverá atender para o credenciamento para licenciar as atividades que estão definidas no anexo único da mesma resolução.
    {municipio} - Res. CEMAm n° 153, de 02 de agosto de 202 - Nível 2.

    Decreto Estadual nº 10.054, de 25 de fevereiro de 2022. Tipologia e porte dos empreendimentos e atividades sujeitos ao licenciamento ambiental.
    Grupo G2: Empreendimentos Urbanísticos
    G2.5 - Uso do solo, em imóveis rurais, decorrente de desmembramento imobiliário, para a formação de chácaras, assentamentos, ecovilas, condomínios, uso por multipropriedades, uso por multirresidências e/ou ocupações de segunda residência ou lazer, observada a legislação de regência sobre a natureza da ocupação em áreas rurais.

    DECRETO nº 9.710, de 03 de setembro de 2020. Que dispõe sobre as normas gerais para o Licenciamento Ambiental no Estado de Goiás e dá outras providências.
    Art. 13. O Poder Público, no exercício de sua competência de controle, expedirá as seguintes licenças:
    VI – Licença Corretiva – LC: ato administrativo que regulariza atividade ou empreendimento em instalação ou operação sem licença ambiental, por meio da fixação de condicionantes que viabilizam sua continuidade em conformidade com as normas ambientais;

    A degradação ambiental do parcelamento do solo é um processo contínuo que se intensifica à medida que a urbanização avança sem planejamento adequado. Com supressões para construções e acessos, há uma destruição imediata de habitats naturais, seguida por impactos persistentes como a erosão do solo e a perda de biodiversidade. A impermeabilização do solo aumenta o escoamento superficial, causando enchentes e assoreamento de corpos d'água, enquanto a poluição contínua de esgotos e resíduos sólidos compromete a qualidade dos recursos hídricos. Esse ciclo de degradação perpetua-se com a pressão crescente sobre infraestruturas e serviços públicos, agravando os problemas ambientais e dificultando a recuperação das áreas afetadas. Sem medidas mitigatórias eficazes, a degradação ambiental torna-se um problema cumulativo, exacerbando os impactos negativos sobre o ecossistema e a qualidade de vida das populações locais.

    A Instrução Normativa n. 50/97 do INCRA, que estabelece as zonas típicas de módulo e estende a fração mínima de parcelamento, estabelece o seguinte:
    Art. 2° - Ficam estendidas a Fração Mínima de Parcelamento - FMP correspondente ao módulo de exploração hortigranjeira das respectivas zonas típicas, prevista para as capitais dos estados, aos municípios classificados nas Zonas Típicas de módulo "B" e "C", de acordo com o parágrafo 2o do artigo 8° da Lei n°5.868, de 12 de dezembro de1972.

    A fração mínima de parcelamento para imóveis rurais (FMP), no Estado de Goiás, está definida na forma abaixo relacionada, com respaldo na Instrução Especial do INCRA n. 50/97, revogada e atualizada pela INSTRUÇÃO ESPECIAL n. 5, de 29 de julho de 2022, levando-se em consideração que o Estado de Goiás é dividido em 18 microrregiões, conforme mapa em anexo, segundo o IBGE, de acordo com a Resolução — PR n° 11 de 05/06/90.

    Infere-se da referida legislação que o imóvel rural, estando localizado na região {microrregiao}, município de {municipio}/GO deve observar uma área mínima de {fracao_minima}, sob pena de nulidade do fracionamento irregular.

    Foram evidenciados por meio dos Contratos de Compra e Venda que os vendedores fracionaram lotes em dimensões bem inferiores ao permitido. É sabido que estes contratos não podem ser averbados no cartório de registro imobiliário, justamente em virtude destas chácaras não respeitarem o tamanho mínimo legal, causando, assim, prejuízo aos compradores.

    Além da não observância ao limite mínimo de parcelamento, os vendedores tampouco providenciaram a regularização do imóvel para este fim.

    {'=' * 60}
    {municipio}, {data_ocorrencia}
    _____________________________________
    Assinatura do Fiscal
    """
        return relatorio

    def gerar_relatorio_parcelamento_ui(self):
        try:
            dados = {}
            print("=" * 60)
            print("COLETANDO DADOS DA ABA PARCELAMENTO:")

            for nome, campo in self.campos_parcelamento.items():
                if hasattr(campo, 'toPlainText'):
                    dados[nome] = campo.toPlainText()
                    print(f"{nome}: {dados[nome][:100]}...")
                elif hasattr(campo, 'currentText'):
                    dados[nome] = campo.currentText()
                    print(f"{nome}: {dados[nome]}")
                else:
                    dados[nome] = campo.text()
                    print(f"{nome}: {dados[nome][:100]}...")

            print("=" * 60)

            if not dados.get('parc_data_ocorrencia'):
                QMessageBox.warning(self, "Aviso", "O campo 'Data da Ocorrência' é obrigatório!")
                return

            relatorio = self.gerar_relatorio_parcelamento(dados)

            pasta = QFileDialog.getExistingDirectory(self, "Escolha a pasta para salvar")
            if not pasta:
                return

            nome_arquivo = f"Relatorio_Parcelamento_{dados.get('imovel', 'NAO_INFORMADO')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            caminho = os.path.join(pasta, nome_arquivo)

            with open(caminho, 'w', encoding='utf-8') as f:
                f.write(relatorio)

            QMessageBox.information(self, "Sucesso", f"Relatório gerado com sucesso!\n{caminho}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar relatório: {str(e)}")
        
    def obter_texto_areas_selecionadas(self):
        """Retorna o texto formatado das áreas selecionadas usando os templates"""
        textos = []
        
        imovel = self.inputs["imovel"].text().strip() if "imovel" in self.inputs and self.inputs["imovel"].text() else "NÃO INFORMADO"
        municipio = self.inputs["municipio"].text().strip() if "municipio" in self.inputs and self.inputs["municipio"].text() else "NÃO INFORMADO"
        uf = self.inputs["uf"].text().strip() if "uf" in self.inputs and self.inputs["uf"].text() else "GO"
        
        tem_area_selecionada = False
        
        for chave in ["app", "rl", "fora", "dano"]:
            check = getattr(self, f"check_{chave}", None)
            if check and check.isChecked():
                tem_area_selecionada = True
                area_data = self.areas_inputs.get(chave, {})
                
                area_widget = area_data.get("area")
                area_text = area_widget.text().strip() if area_widget and hasattr(area_widget, 'text') else "0"
                
                valor_widget = area_data.get("valor")
                valor_text = valor_widget.text().strip() if valor_widget and hasattr(valor_widget, 'text') else "NÃO CALCULADO"
                
                templates = self.obter_template_texto_area(chave)
                
                # Formata o texto do Auto
                texto_auto = templates["auto"].format(
                    area=area_text,
                    imovel=imovel,
                    municipio=municipio,
                    uf=uf,
                    valor_multa=valor_text
                )
                
                texto = f"""
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ {self.obter_nome_area(chave)}                              │
    └─────────────────────────────────────────────────────────────────────────────┘

    📄 AUTO DE INFRAÇÃO:
    {texto_auto}

    """
                # Adiciona embargo se não for dano
                if chave != "dano" and templates["embargo"]:
                    texto_embargo = templates["embargo"].format(
                        area=area_text,
                        imovel=imovel,
                        municipio=municipio,
                        uf=uf
                    )
                    texto += f"""
    🚫 EMBARGO:
    {texto_embargo}

    """
                textos.append(texto)
        
        if not tem_area_selecionada:
            return """
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ NENHUMA ÁREA SELECIONADA                                                     │
    └─────────────────────────────────────────────────────────────────────────────┘

    Não há áreas identificadas para este processo de fiscalização.
    """
        
        return "\n".join(textos)

    def obter_descricao_infracao(self):
        """Retorna a descrição das infrações com títulos claros e descrição de acesso"""
        descricoes = []

        imovel = self.inputs["imovel"].text().strip() if "imovel" in self.inputs and self.inputs[
            "imovel"].text() else "NÃO INFORMADO"
        municipio = self.inputs["municipio"].text().strip() if "municipio" in self.inputs and self.inputs[
            "municipio"].text() else "NÃO INFORMADO"
        uf = self.inputs["uf"].text().strip() if "uf" in self.inputs and self.inputs["uf"].text() else "GO"

        descricao_acesso = self.obs.toPlainText().strip() if hasattr(self, 'obs') else ""

        tem_infracao = False

        # ==========================================================
        # 1. ÁREAS (APP, RL, Fora, Dano) - COM TÍTULOS
        # ==========================================================
        for chave in ["app", "rl", "fora", "dano"]:
            check = getattr(self, f"check_{chave}", None)
            if check and check.isChecked():
                tem_infracao = True
                area_data = self.areas_inputs.get(chave, {})

                area_widget = area_data.get("area")
                area_text = area_widget.text().strip() if area_widget and hasattr(area_widget, 'text') else "0"

                valor_widget = area_data.get("valor")
                valor_text = valor_widget.text().strip() if valor_widget and hasattr(valor_widget,
                                                                                     'text') else "NÃO CALCULADO"

                templates = self.obter_template_texto_area(chave)

                nomes_com_artigo = {
                    "app": "Art. 43 - Área de Preservação Permanente (APP)",
                    "rl": "Art. 51 - Reserva Legal (RL)",
                    "fora": "Art. 52 - Área Fora de APP e RL (Área Passível)",
                    "dano": "Art. 53 - Dano Ambiental (Corte de Árvores Isoladas)"
                }
                titulo_area = nomes_com_artigo.get(chave, chave.upper())

                texto_auto = templates["auto"].format(
                    area=area_text,
                    imovel=imovel,
                    municipio=municipio,
                    uf=uf,
                    valor_multa=valor_text
                )

                linhas = texto_auto.split('\n')
                if len(linhas) > 1:
                    texto_auto = '\n'.join(linhas[1:]).strip()

                descricao = f"""
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ 📌 {titulo_area}
    └─────────────────────────────────────────────────────────────────────────────┘

    📄 AUTO DE INFRAÇÃO:
    {texto_auto}"""

                if chave != "dano" and templates["embargo"]:
                    texto_embargo = templates["embargo"].format(
                        area=area_text,
                        imovel=imovel,
                        municipio=municipio,
                        uf=uf
                    )
                    linhas_embargo = texto_embargo.split('\n')
                    if len(linhas_embargo) > 1:
                        texto_embargo = '\n'.join(linhas_embargo[1:]).strip()
                    descricao += f"""

    🚫 EMBARGO:
    {texto_embargo}"""

                if descricao_acesso:
                    descricao += f"""

    📍 DESCRIÇÃO DE ACESSO AO IMÓVEL:
    {descricao_acesso}"""

                descricoes.append(descricao)

        # ==========================================================
        # 2. ARTIGOS (48, 66, 79) - COM TÍTULOS
        # ==========================================================
        artigos_config = {
            "art48": {
                "titulo": "Art. 48 - Impedir Regeneração de Vegetação",
                "template_auto": """Por impedir ou dificultar a regeneração natural de florestas e demais formas de vegetação nativa, na área de {area} hectares, sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.""",
                "template_embargo": """Fica embargada a área de {area} hectares, por impedir ou dificultar a regeneração natural de florestas e demais formas de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "art66": {
                "titulo": "Art. 66 - Executar Atividade sem Licença",
                "template_auto": """Por construir, reformar, ampliar, instalar ou fazer funcionar estabelecimentos, atividades, obras ou serviços utilizadores de recursos ambientais, considerados efetiva ou potencialmente poluidores, sem licença ou autorização dos órgãos ambientais competentes, na área de {area} hectares, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.""",
                "template_embargo": """Fica embargada a área de {area} hectares, por executar atividade sem licença ou autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "art79": {
                "titulo": "Art. 79 - Descumprimento de Embargo",
                "template_auto": """Por descumprir embargo instituído por autoridade ambiental competente, reincidindo na atividade embargada, na área de {area} hectares, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}."""
            }
        }

        for chave, config in artigos_config.items():
            check = getattr(self, f"check_{chave}", None)
            if check and check.isChecked():
                tem_infracao = True

                artigo_data = self.artigos_inputs.get(chave, {})

                area_widget = artigo_data.get("area")
                area_text = area_widget.text().strip() if area_widget and hasattr(area_widget, 'text') else "0"

                valor_widget = artigo_data.get("valor")
                valor_text = valor_widget.text().strip() if valor_widget and hasattr(valor_widget,
                                                                                     'text') else "NÃO CALCULADO"

                texto_auto = config['template_auto'].format(
                    area=area_text,
                    imovel=imovel,
                    municipio=municipio,
                    uf=uf,
                    valor_multa=valor_text
                )

                descricao = f"""
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ 📌 {config['titulo']}
    └─────────────────────────────────────────────────────────────────────────────┘

    📄 AUTO DE INFRAÇÃO:
    {texto_auto}"""

                if 'template_embargo' in config:
                    texto_embargo = config['template_embargo'].format(
                        area=area_text,
                        imovel=imovel,
                        municipio=municipio,
                        uf=uf
                    )
                    descricao += f"""

    🚫 EMBARGO:
    {texto_embargo}"""

                if descricao_acesso:
                    descricao += f"""

    📍 DESCRIÇÃO DE ACESSO AO IMÓVEL:
    {descricao_acesso}"""

                descricoes.append(descricao)

        # ==========================================================
        # 3. ARTIGO GENÉRICO
        # ==========================================================
        if hasattr(self, 'check_art_generico') and self.check_art_generico and self.check_art_generico.isChecked():
            art_data = self.artigos_inputs.get("art_generico", {})
            if art_data:
                tem_infracao = True
                titulo = art_data.get("titulo", "Art. Genérico")
                area_text = art_data.get("area", QLineEdit()).text().strip() or "0"
                valor_text = art_data.get("valor", QLineEdit()).text().strip() or "NÃO CALCULADO"

                # Pega a descrição da infração que o usuário digitou
                desc_infracao_widget = art_data.get("descricao_infracao", None)
                if desc_infracao_widget and hasattr(desc_infracao_widget, 'toPlainText'):
                    descricao_infracao_text = desc_infracao_widget.toPlainText().strip()
                else:
                    descricao_infracao_text = "cometer a infração"

                if not descricao_infracao_text:
                    descricao_infracao_text = "cometer a infração"

                texto_auto = f"""Por {descricao_infracao_text}, na área de {area_text} hectares, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_text}."""

                descricao = f"""
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ 📌 {titulo}
    └─────────────────────────────────────────────────────────────────────────────┘

    📄 AUTO DE INFRAÇÃO:
    {texto_auto}

    🚫 EMBARGO:
    Fica embargada a área de {area_text} hectares, por {descricao_infracao_text}, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""

                if descricao_acesso:
                    descricao += f"""

    📍 DESCRIÇÃO DE ACESSO AO IMÓVEL:
    {descricao_acesso}"""

                descricoes.append(descricao)

        # ==========================================================
        # RETORNO
        # ==========================================================
        if not tem_infracao:
            return """
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ 📌 NENHUMA INFRAÇÃO IDENTIFICADA
    └─────────────────────────────────────────────────────────────────────────────┘

    Após análise detalhada das imagens de satélite e dados geoespaciais, não foram 
    identificadas irregularidades ambientais no imóvel fiscalizado neste período.
    """

        if not descricoes:
            return """
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ 📌 NENHUMA INFRAÇÃO IDENTIFICADA
    └─────────────────────────────────────────────────────────────────────────────┘

    Não há infrações selecionadas para este processo de fiscalização.
    """

        return "\n" + "\n\n".join(descricoes)

    def _gerar_descricao_infracao_por_tipo(self, tipo, area, imovel, municipio, uf):
        """Gera descrição específica por tipo de infração"""
        descricoes = {
            "app": f"Constata-se a supressão de {area} hectares de vegetação nativa em Área de "
                f"Preservação Permanente (APP), sem a devida autorização do órgão ambiental "
                f"competente. Tal conduta viola o disposto na Lei Federal nº 12.651/2012 "
                f"(Código Florestal) e no Decreto Federal nº 6.514/2008, caracterizando "
                f"infração ambiental de natureza grave.",
            
            "rl": f"Verifica-se a supressão de {area} hectares de vegetação nativa em área de "
                f"Reserva Legal (RL), sem autorização prévia do órgão ambiental competente. "
                f"A conduta infringe o estabelecido no Código Florestal Brasileiro e demais "
                f"normas ambientais vigentes.",
            
            "fora": f"Identifica-se a supressão de {area} hectares de vegetação nativa em área "
                    f"fora de Reserva Legal e fora de Área de Preservação Permanente, sem a "
                    f"respectiva autorização ambiental. A atividade configura infração ambiental "
                    f"passível de autuação e embargo.",
            
            "dano": f"Caracteriza-se dano ambiental em {area} hectares de vegetação nativa, "
                    f"decorrente de corte de árvores isoladas sem autorização legal. O dano "
                    f"ambiental causado compromete a integridade do ecossistema local."
        }
        return descricoes.get(tipo, descricoes["fora"])
        
    def obter_nome_imovel(self):
        """Retorna o nome do imóvel dos dados gerais"""
        return self.inputs["imovel"].text().strip() if "imovel" in self.inputs and self.inputs["imovel"].text() else "NÃO INFORMADO"

    def obter_municipio(self):
        """Retorna o município dos dados gerais"""
        return self.inputs["municipio"].text().strip() if "municipio" in self.inputs and self.inputs["municipio"].text() else "NÃO INFORMADO"

        
    def obter_embargo(self):
        """Retorna o texto do embargo da primeira área selecionada"""
        if self.check_app and self.check_app.isChecked():
            widget = self.areas_inputs.get("app", {}).get("embargo")
            if widget and hasattr(widget, 'toPlainText'):
                return widget.toPlainText().strip() or "NÃO INFORMADO"
            return "NÃO INFORMADO"
        elif self.check_rl and self.check_rl.isChecked():
            widget = self.areas_inputs.get("rl", {}).get("embargo")
            if widget and hasattr(widget, 'toPlainText'):
                return widget.toPlainText().strip() or "NÃO INFORMADO"
            return "NÃO INFORMADO"
        elif self.check_fora and self.check_fora.isChecked():
            widget = self.areas_inputs.get("fora", {}).get("embargo")
            if widget and hasattr(widget, 'toPlainText'):
                return widget.toPlainText().strip() or "NÃO INFORMADO"
            return "NÃO INFORMADO"
        return "NÃO INFORMADO"  # Dano não tem embargo
    
    def obter_auto_infracao(self):
        """Retorna o texto do auto de infração da primeira área selecionada usando template"""
        imovel = self.inputs["imovel"].text().strip() if "imovel" in self.inputs and self.inputs["imovel"].text() else "NÃO INFORMADO"
        municipio = self.inputs["municipio"].text().strip() if "municipio" in self.inputs and self.inputs["municipio"].text() else "NÃO INFORMADO"
        uf = self.inputs["uf"].text().strip() if "uf" in self.inputs and self.inputs["uf"].text() else "GO"
        
        for chave in ["app", "rl", "fora", "dano"]:
            check = getattr(self, f"check_{chave}", None)
            if check and check.isChecked():
                area_data = self.areas_inputs.get(chave, {})
                area_widget = area_data.get("area")
                area_text = area_widget.text().strip() if area_widget and hasattr(area_widget, 'text') else "0"
                
                valor_widget = area_data.get("valor")
                valor_text = valor_widget.text().strip() if valor_widget and hasattr(valor_widget, 'text') else "NÃO CALCULADO"
                
                templates = self.obter_template_texto_area(chave)
                
                return templates["auto"].format(
                    area=area_text,
                    imovel=imovel,
                    municipio=municipio,
                    uf=uf,
                    valor_multa=valor_text
                )
        return "NENHUM AUTO GERADO"

    def obter_template_texto_area(self, chave):
        """Retorna o template de texto da área específica"""
        templates = {
            "app": {
                "auto": """SUPRESSÃO EM ÁREA DE PRESERVAÇÃO PERMANENTE

        Por suprimir {area} hectares de vegetação nativa em Área de Preservação Permanente (APP), sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}.
        Valor da Multa em APP: {valor_multa}""",
                "embargo": """EMBARGO - ÁREA DE PRESERVAÇÃO PERMANENTE

        Fica embargada a área de {area} hectares em Área de Preservação Permanente (APP), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "rl": {
                "auto": """SUPRESSÃO EM RESERVA LEGAL

        Por suprimir {area} hectares de vegetação nativa em área de Reserva Legal (RL), sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}.
        Valor da Multa em RL: {valor_multa}""",
                "embargo": """EMBARGO - RESERVA LEGAL

        Fica embargada a área de {area} hectares em área de Reserva Legal (RL), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "fora": {
                "auto": """SUPRESSÃO EM ÁREA PASSÍVEL

        Por suprimir {area} hectares de vegetação nativa em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}.
        Valor da Multa em Área Passível: {valor_multa}""",
                "embargo": """EMBARGO - ÁREA PASSÍVEL

        Fica embargada a área de {area} hectares em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "dano": {
                "auto": """DANO AMBIENTAL - CORTE DE ÁRVORES ISOLADAS

        Por danificar {area} hectares de vegetação nativa em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}. Esta área não fica embargada.
        Valor da Multa em Área Danificada: {valor_multa}""",
                "embargo": ""
            }
        }
        return templates.get(chave, templates["fora"])

    def obter_nome_area(self, chave):
        """Retorna o nome da área para exibição"""
        nomes = {
            "app": "ÁREA DE PRESERVAÇÃO PERMANENTE (APP)",
            "rl": "RESERVA LEGAL (RL)",
            "fora": "ÁREA FORA DE APP E RL",
            "dano": "DANO AMBIENTAL"
        }
        return nomes.get(chave, "ÁREA NÃO IDENTIFICADA")
    
    def is_uc_marcado(self):
        """Verifica se o checkbox de UC está marcado (em qualquer uma das abas)"""
        marcado_areas = hasattr(self, 'check_uc') and self.check_uc and self.check_uc.isChecked()
        marcado_artigos = hasattr(self, 'check_uc_artigos') and self.check_uc_artigos and self.check_uc_artigos.isChecked()
        return bool(marcado_areas or marcado_artigos)

    def sincronizar_checkbox_uc(self, origem, state):
        """Mantém os checkboxes de UC das abas 'Autos e Embargos (Supressão)' e
        'Autos e Embargos (Demais)' sincronizados e recalcula as multas."""
        estado_marcado = (state == Qt.Checked) or (origem is not None and origem.isChecked())

        for cb in [getattr(self, 'check_uc', None), getattr(self, 'check_uc_artigos', None)]:
            if cb and cb is not origem and cb.isChecked() != estado_marcado:
                cb.blockSignals(True)
                cb.setChecked(estado_marcado)
                cb.blockSignals(False)

        self.on_uc_changed(Qt.Checked if estado_marcado else Qt.Unchecked)

    def adicionar_tooltips(self):
        """Adiciona tooltips explicativas nos campos principais"""
        tooltips = {
            "processo": "Número do processo administrativo SEMAD",
            "imovel": "Nome completo do imóvel rural conforme CAR",
            "car": "Número do Cadastro Ambiental Rural (formato: GO-XXXXXX-XXXXXX)",
            "alertas": "Números dos alertas do Mapbiomas separados por vírgula",
            "intervalo_supressao": "Período em que ocorreu a supressão (ex: janeiro a abril de 2025)"
        }
        
        for campo, texto in tooltips.items():
            if campo in self.inputs:
                self.inputs[campo].setToolTip(texto)
                
    def validar_area(self, area_input):   # <--- COLOQUE AQUI
        """Valida e formata o campo de área"""
        texto = area_input.text().strip()
        if not texto:
            return None
        
        # Substitui vírgula por ponto
        texto = texto.replace(",", ".")
        
        # Remove múltiplos pontos
        partes = texto.split(".")
        if len(partes) > 2:
            texto = partes[0] + "." + "".join(partes[1:])
        
        try:
            valor = float(texto)
            if valor < 0:
                QMessageBox.warning(self, "Erro", "A área não pode ser negativa!")
                return None
            return valor
        except ValueError:
            QMessageBox.warning(self, "Erro", f"Valor inválido: '{texto}'. Use formato como 5,02 ou 5.02")
            return None

    def preencher_texto_supressao(self, state, tipo, area_input, auto_input, embargo_input, valor_input, texto_template):
        """Preenche os campos de Auto e Embargo com o texto padrão quando o checkbox é marcado"""
        if state == Qt.Checked:  # Checkbox marcado
            # Obter os valores atuais
            area_texto = area_input.text().strip()
            valor_texto = valor_input.text().strip()
            
            # Obter dados do imóvel dos campos principais
            imovel = self.inputs["imovel"].text().strip() if "imovel" in self.inputs else ""
            municipio = self.inputs["municipio"].text().strip() if "municipio" in self.inputs else ""
            
            # Se não tiver imóvel ou município, usar valores padrão
            if not imovel:
                imovel = "NÃO INFORMADO"
            if not municipio:
                municipio = "NÃO INFORMADO"
            
            uf = self.inputs["uf"].text().strip() if "uf" in self.inputs and self.inputs["uf"].text() else "GO"
            
            # Se a área não foi preenchida, usar um placeholder
            if not area_texto:
                area_texto = "[Área a ser preenchida]"
            
            # Se o valor da multa não foi calculado, usar placeholder
            if not valor_texto:
                valor_texto = "[Valor a ser calculado]"
            
            # Substituir placeholders no texto do Auto
            texto_formatado = texto_template.format(
                area=area_texto,
                imovel=imovel,
                municipio=municipio,
                uf=uf,
                valor_multa=valor_texto
            )
            
            # Preencher o campo de Auto
            auto_input.setText(texto_formatado)
            
            # Preencher o campo de Embargo (se existir e não for dano)
            if embargo_input:
                # Para o embargo, tenta formatar sem valor_multa
                try:
                    texto_embargo_formatado = texto_template.format(
                        area=area_texto,
                        imovel=imovel,
                        municipio=municipio,
                        uf=uf
                    )
                except KeyError:
                    # Se o template espera valor_multa, usa uma versão sem
                    texto_embargo_formatado = texto_template.split('Valor da Multa')[0].format(
                        area=area_texto,
                        imovel=imovel,
                        municipio=municipio,
                        uf=uf
                    )
                embargo_input.setText(texto_embargo_formatado)

    def gerar_despacho_autocomposicao(self, dados):
        """Gera o conteúdo do Despacho de Autocomposição"""
        # Usa .get() com valor padrão para evitar KeyError
        areas_selecionadas = self.obter_texto_areas_selecionadas()
        descricao_infracao = self.obter_descricao_infracao()

        return f"""DESPACHO DE AUTOCOMPOSIÇÃO
    {'=' * 60}

    Processo Administrativo nº: {dados.get('processo', 'NÃO INFORMADO')}
    Data: {dados.get('data', 'NÃO INFORMADO')}
    OS: {dados.get('os', 'NÃO INFORMADO')}

    I - RELATÓRIO
    Trata-se de processo administrativo instaurado em face do proprietário do imóvel rural 
    denominado {dados.get('imovel', 'NÃO INFORMADO')}, localizado no município de {dados.get('municipio', 'NÃO INFORMADO')} - {dados.get('uf', 'GO')}, 
    de propriedade de {dados.get('proprietario', 'NÃO INFORMADO')}, inscrito no CPF/CNPJ sob o nº {dados.get('cpf', 'NÃO INFORMADO')}.

    II - DAS INFRAÇÕES IDENTIFICADAS
    Foram identificadas as seguintes infrações ambientais no imóvel fiscalizado:

    ÁREAS IDENTIFICADAS:
    {areas_selecionadas}

    DESCRIÇÃO DAS INFRAÇÕES:
    {descricao_infracao}

    III - DA AUTOCOMPOSIÇÃO
    O infrator manifestou interesse em celebrar Termo de Autocomposição, comprometendo-se a:

    1. Recuperar as áreas degradadas no prazo estabelecido;
    2. Pagar a multa no valor estipulado no Auto de Infração;
    3. Paralisar imediatamente as atividades irregulares;
    4. Apresentar relatórios periódicos de recuperação ambiental.

    IV - DAS CONDIÇÕES
    O cumprimento das obrigações assumidas será acompanhado por esta SEMAD, podendo ser 
    requisitadas informações complementares a qualquer tempo.

    V - DISPOSITIVO
    Diante do exposto, HOMOLOGO o pedido de autocomposição e determino o sobrestamento do 
    processo pelo prazo de 90 (noventa) dias para cumprimento das obrigações assumidas.

    Ciência ao Ministério Público.

    {dados.get('municipio', 'NÃO INFORMADO')}, {dados.get('data', 'NÃO INFORMADO')}
    _____________________________________
    Assinatura do Presidente da Comissão
    """
    def gerar_modelo_minuta(self, dados):
        """Gera o conteúdo do Modelo de Minuta de Auto de Infração"""
        return f"""MODELO DE MINUTA DE AUTO DE INFRAÇÃO

    Processo Administrativo nº: {dados.get('processo', 'NÃO INFORMADO')}
    Data: {dados.get('data', 'NÃO INFORMADO')}

    AUTO DE INFRAÇÃO Nº: _______________

    Aos {dados.get('data', 'NÃO INFORMADO')}, o fiscal ambiental da SEMAD, no exercício de suas atribuições, 
    lavrou o presente Auto de Infração em face de:

    INFRATOR(A): {dados.get('proprietario', 'NÃO INFORMADO')}
    CPF/CNPJ: {dados.get('cpf', 'NÃO INFORMADO')}
    ENDEREÇO: _______________________________

    IMÓVEL RURAL: {dados.get('imovel', 'NÃO INFORMADO')}
    MUNICÍPIO: {dados.get('municipio', 'NÃO INFORMADO')}
    UF: {dados.get('uf', 'GO')}
    CAR: {dados.get('car', 'NÃO INFORMADO')}
    COORDENADAS: {dados.get('coordenadas', 'NÃO INFORMADO')}

    DESCRIÇÃO DA INFRAÇÃO:
    {dados.get('observacoes', 'NÃO INFORMADO')}

    ARTIGO(S) / INFRAÇÃO(ÕES) IDENTIFICADA(S):
    {self.obter_descricao_infracao()}

    ENCAMINHAMENTO:
    O presente Auto de Infração será encaminhado para julgamento.

    ASSINATURA DO FISCAL:
    _____________________________________
    Data: {dados.get('data', 'NÃO INFORMADO')}
    """

    def gerar_relatorio_autodenuncia(self, dados):
        """Gera o conteúdo do Relatório de Autodenúncia"""
        return f"""RELATÓRIO DE AUTODENÚNCIA

    Processo: {dados.get('processo', 'NÃO INFORMADO')}
    Data: {dados.get('data', 'NÃO INFORMADO')}

    I - IDENTIFICAÇÃO DO AUTODENUNCIANTE
    Nome: {dados.get('proprietario', 'NÃO INFORMADO')}
    CPF/CNPJ: {dados.get('cpf', 'NÃO INFORMADO')}
    Endereço: ___________________________________

    II - IDENTIFICAÇÃO DO IMÓVEL
    Imóvel: {dados.get('imovel', 'NÃO INFORMADO')}
    Município: {dados.get('municipio', 'NÃO INFORMADO')}
    UF: {dados.get('uf', 'GO')}
    CAR: {dados.get('car', 'NÃO INFORMADO')}
    Coordenadas: {dados.get('coordenadas', 'NÃO INFORMADO')}

    III - RELATO DA AUTODENÚNCIA
    {dados.get('observacoes', 'NÃO INFORMADO')}

    IV - INFRAÇÕES IDENTIFICADAS
    {self.obter_descricao_infracao()}

    V - PROVIDÊNCIAS ADOTADAS
    1. Registro da autodenúncia;
    2. Análise preliminar das informações;
    3. Agendamento de fiscalização in loco.

    VI - ANÁLISE DA EQUIPE TÉCNICA
    A equipe técnica analisou as informações prestadas e constatou que:

    [Análise a ser preenchida pela equipe]

    VII - CONCLUSÃO
    Diante do exposto, recomenda-se:

    [Recomendações a serem preenchidas]

    _____________________________________
    Assinatura do Fiscal
    Data: {dados.get('data', 'NÃO INFORMADO')}
    """           
    def criar_checkboxes_areas(self):
        """Cria os checkboxes de áreas identificadas com controle de ativação"""
        grupo_check = QGroupBox("📍 Áreas Identificadas")
        grupo_check.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 2px solid #006b3f;
                border-radius: 10px;
                padding: 10px;
            }
            QGroupBox::title {
                color: #006b3f;
                font-weight: bold;
                font-size: 12pt;
                padding: 0 10px;
            }
        """)
        
        check_layout = QHBoxLayout()
        check_layout.setSpacing(20)
        check_layout.setContentsMargins(15, 10, 15, 10)
        
        # Criar checkboxes com IDs para referência
        self.check_app = QCheckBox("🌳 APP")
        self.check_app.setObjectName("check_app")
        
        self.check_rl = QCheckBox("🌿 Reserva Legal")
        self.check_rl.setObjectName("check_rl")
        
        self.check_fora = QCheckBox("📐 Fora de APP e RL")
        self.check_fora.setObjectName("check_fora")
        
        self.check_dano = QCheckBox("⚠️ Dano Ambiental")
        self.check_dano.setObjectName("check_dano")
        
        self.check_uc = QCheckBox("🏛️ UC - Unidade de Conservação")
        self.check_uc.setObjectName("uc-check")
        
        # Conectar os checkboxes aos grupos correspondentes
        self.check_app.toggled.connect(lambda checked: self.toggle_area_group("app", checked))
        self.check_rl.toggled.connect(lambda checked: self.toggle_area_group("rl", checked))
        self.check_fora.toggled.connect(lambda checked: self.toggle_area_group("fora", checked))
        self.check_dano.toggled.connect(lambda checked: self.toggle_area_group("dano", checked))
        
        # Conexão do UC (mantém sincronia com a outra aba)
        self.check_uc.stateChanged.connect(lambda state: self.sincronizar_checkbox_uc(self.check_uc, state))
        
        # Estilização dos checkboxes
        for cb in [self.check_app, self.check_rl, self.check_fora, self.check_dano, self.check_uc]:
            cb.setStyleSheet("""
                QCheckBox {
                    spacing: 10px;
                    font-weight: 500;
                    padding: 4px 8px;
                    border-radius: 6px;
                }
                QCheckBox:hover {
                    background-color: #f0f2f5;
                }
            """)
        
        check_layout.addWidget(self.check_app)
        check_layout.addWidget(self.check_rl)
        check_layout.addWidget(self.check_fora)
        check_layout.addWidget(self.check_dano)
        check_layout.addWidget(self.check_uc)
        check_layout.addStretch()
        
        grupo_check.setLayout(check_layout)
        return grupo_check

    def toggle_area_group(self, chave, ativo):
        """Ativa ou desativa o grupo de área correspondente"""
        if chave in self.areas_inputs:
            grupo = self.areas_inputs[chave].get("grupo")
            if grupo:
                grupo.setEnabled(ativo)
                # Estiliza visualmente quando desativado
                if ativo:
                    grupo.setStyleSheet(grupo.styleSheet() + """
                        QGroupBox { opacity: 1; }
                        QGroupBox::title { color: #006b3f; }
                    """)
                else:
                    grupo.setStyleSheet(grupo.styleSheet() + """
                        QGroupBox { opacity: 0.5; background-color: #f8f9fa; }
                        QGroupBox::title { color: #9aa0a6; }
                    """)
            
            # Limpa os campos se desativar
            if not ativo:
                dados = self.areas_inputs[chave]
                if "area" in dados and dados["area"]:
                    dados["area"].clear()
                if "auto" in dados and dados["auto"]:
                    dados["auto"].clear()
                if "embargo" in dados and dados["embargo"]:
                    dados["embargo"].clear()
                if "valor" in dados and dados["valor"]:
                    dados["valor"].clear()
                    dados["valor"].setPlaceholderText("Área não selecionada")
        
        # Atualiza a pré-visualização
        self.atualizar_preview()

        def criar_grupo_area(self, chave, dados_config):
            """Cria um grupo de área com controle de habilitação"""
            grupo = QGroupBox(dados_config["titulo"])
            grupo.setObjectName(f"grupo_{chave}")
            grupo.setProperty("class", "area-card")
            
            # Começa desabilitado por padrão (só habilita se o checkbox estiver marcado)
            check = getattr(self, f"check_{chave}", None)
            if check:
                grupo.setEnabled(check.isChecked())
            else:
                grupo.setEnabled(False)
            
            grupo.setStyleSheet(f"""
                QGroupBox[class="area-card"] {{
                    background-color: #fafbfc;
                    border: 1px solid #e8ecf0;
                    border-radius: 8px;
                    margin-top: 8px;
                    padding-top: 12px;
                }}
                QGroupBox[class="area-card"]::title {{
                    color: #006b3f;
                    font-weight: 600;
                    font-size: 10pt;
                    padding: 0 8px;
                }}
                QGroupBox:disabled {{
                    opacity: 0.5;
                    background-color: #f8f9fa;
                }}
                QGroupBox:disabled::title {{
                    color: #9aa0a6;
                }}
            """)
    
    def atualizar_descricao_acesso(self):
        """Atualiza automaticamente a descrição de acesso ao imóvel
        com base nos campos Imóvel, Município e Coordenadas."""
        if not hasattr(self, 'obs'):
            return
        imovel = self.inputs.get("imovel", QLineEdit()).text().strip() or "NÃO INFORMADO"
        municipio = self.inputs.get("municipio", QLineEdit()).text().strip() or "NÃO INFORMADO"
        coordenadas = self.inputs.get("coordenadas", QLineEdit()).text().strip() or "NÃO INFORMADAS"

        texto = (
            f"Trata-se da propriedade {imovel}, situada no município de {municipio}, "
            f"com coordenadas geográficas {coordenadas}, referenciadas ao sistema geodésico SIRGAS 2000."
        )
        self.obs.setText(texto)
        
    def criar_aba_dados_gerais(self):
        widget = QWidget()
        widget.setStyleSheet("background-color: #F5F7F8;")
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # ==========================================================
        # CARD - INFORMAÇÕES DO PROCESSO (grid de 2 colunas)
        # ==========================================================
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e1e4e8;
                border-radius: 10px;
            }
        """)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(20, 18, 20, 20)
        card_layout.setSpacing(14)

        titulo_card = QLabel("📋 Informações do Processo")
        titulo_card.setStyleSheet("""
            font-weight: 700;
            font-size: 13pt;
            color: #0B6B3A;
            padding-bottom: 8px;
            border-bottom: 1px solid #e1e4e8;
        """)
        card_layout.addWidget(titulo_card)

        grid = QGridLayout()
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(4)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)

        campos = [
            ("Processo", "processo"),
            ("Nome do Imóvel", "imovel"),
            ("Município", "municipio"),
            ("CAR", "car"),
            ("Proprietário", "proprietario"),
            ("CPF/CNPJ", "cpf"),
            ("Coordenadas", "coordenadas"),
            ("UF", "uf"),
            ("OS (Ordem de serviço)", "os"),
            ("Alertas Mapbiomas", "alertas"),
            ("Intervalo Supressão Total", "intervalo_supressao"),
        ]

        estilo_label = """
            font-weight: 500;
            color: #24292e;
            font-size: 10pt;
            background: transparent;
            border: none;
            padding: 0px;
        """
        estilo_input = """
            QLineEdit {
                border: 1px solid #DDE3E8;
                border-radius: 6px;
                padding: 8px 10px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1.5px solid #0B6B3A;
            }
        """

        for i, (texto, nome) in enumerate(campos):
            linha = i // 2
            coluna = i % 2

            campo_box = QVBoxLayout()
            campo_box.setSpacing(4)

            label = QLabel(texto)
            label.setStyleSheet(estilo_label)
            campo_box.addWidget(label)

            entrada = QLineEdit()
            entrada.setStyleSheet(estilo_input)
            if nome == "uf":
                entrada.setText("GO")
            campo_box.addWidget(entrada)

            self.inputs[nome] = entrada

            wrapper = QWidget()
            wrapper.setLayout(campo_box)
            grid.addWidget(wrapper, linha, coluna)

        card_layout.addLayout(grid)
        layout.addWidget(card)

        # ==========================================================
        # CARD - DESCRIÇÃO DE ACESSO AO IMÓVEL
        # ==========================================================
        card_obs = QFrame()
        card_obs.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e1e4e8;
                border-radius: 10px;
            }
        """)
        obs_layout = QVBoxLayout(card_obs)
        obs_layout.setContentsMargins(20, 18, 20, 20)
        obs_layout.setSpacing(10)

        titulo_obs = QLabel("📍 Descrição de acesso ao imóvel")
        titulo_obs.setStyleSheet("""
            font-weight: 700;
            font-size: 13pt;
            color: #0B6B3A;
            padding-bottom: 8px;
            border-bottom: 1px solid #e1e4e8;
        """)
        obs_layout.addWidget(titulo_obs)

        self.obs = QTextEdit()
        self.obs.setStyleSheet("""
            QTextEdit {
                border: 1px solid #DDE3E8;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        self.obs.setText(
            "Durante análise técnica e fiscalização remota, "
            "foram identificados indícios de supressão de "
            "vegetação nativa."
        )
        obs_layout.addWidget(self.obs)
        layout.addWidget(card_obs)

        # Conecta os campos para atualizar a descrição de acesso automaticamente
        self.inputs["imovel"].textChanged.connect(self.atualizar_descricao_acesso)
        self.inputs["municipio"].textChanged.connect(self.atualizar_descricao_acesso)
        self.inputs["coordenadas"].textChanged.connect(self.atualizar_descricao_acesso)

        layout.addStretch()
        container.setLayout(layout)
        scroll.setWidget(container)

        final_layout = QVBoxLayout()
        final_layout.setContentsMargins(0, 0, 0, 0)
        final_layout.addWidget(scroll)
        widget.setLayout(final_layout)

        return widget

    def criar_aba_areas(self):
        widget = QWidget()
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # Primeiro, criar os checkboxes (antes de usá-los)
        grupo_check = self.criar_checkboxes_areas()
        layout.addWidget(grupo_check)   # ← chips ficam no topo da aba
                
        configuracoes = {
            "app": {
                "titulo": "Art.43 - Área de Preservação Permanente (APP)",
                "descricao": "Valor base: R$ 5.000,00/ha | Teto: R$ 50.000,00/ha",
                "cor": "#f8f9fa",
                "texto_auto": """SUPRESSÃO EM ÁREA DE PRESERVAÇÃO PERMANENTE

    Por suprimir {area} hectares de vegetação nativa em Área de Preservação Permanente (APP), sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.
    Valor da Multa em APP: {valor_multa}""",
                "texto_embargo": """EMBARGO - ÁREA DE PRESERVAÇÃO PERMANENTE

    Fica embargada a área de {area} hectares em Área de Preservação Permanente (APP), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf},."""
            },
            "rl": {
                "titulo": "Reserva Legal (RL)",
                "descricao": "Valor fixo: R$ 5.000,00 por hectare",
                "cor": "#f8f9fa",
                "texto_auto": """SUPRESSÃO EM RESERVA LEGAL

    Por suprimir {area} hectares de vegetação nativa em área de Reserva Legal (RL), sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.
    Valor da Multa em RL: {valor_multa}""",
                "texto_embargo": """EMBARGO - RESERVA LEGAL

    Fica embargada a área de {area} hectares em área de Reserva Legal (RL), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "fora": {
                "titulo": "Fora de APP e de RL (Área Passível)",
                "descricao": "Valor fixo: R$ 1.000,00 por hectare",
                "cor": "#f8f9fa",
                "texto_auto": """SUPRESSÃO EM ÁREA PASSÍVEL

    Por suprimir {area} hectares de vegetação nativa em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.
    Valor da Multa em Área Passível: {valor_multa}""",
                "texto_embargo": """EMBARGO - ÁREA PASSÍVEL

    Fica embargada a área de {area} hectares em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "dano": {
                "titulo": "Dano Ambiental (Corte de árvores isoladas)",
                "descricao": "Valor fixo: R$ 300,00 por hectare",
                "cor": "#f8f9fa",
                "texto_auto": """DANO AMBIENTAL - CORTE DE ÁRVORES ISOLADAS

    Por danificar {area} hectares de vegetação nativa em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}. Esta área não fica embargada.
    Valor da Multa em Área Danificada: {valor_multa}""",
                "texto_embargo": ""  # Dano não tem embargo
            }
        }

        for chave, dados in configuracoes.items():
            grupo = QGroupBox(dados["titulo"])
            grupo.setStyleSheet(f"""
                QGroupBox {{
                    background-color: {dados['cor']};
                    border: 1px solid #e1e4e8;
                    border-radius: 8px;
                    margin-top: 10px;
                    padding-top: 15px;
                }}
                QGroupBox::title {{
                    font-weight: 600;
                    color: #006b3f;
                }}
            """)
            
            main_layout = QVBoxLayout()
            main_layout.setSpacing(10)
            main_layout.setContentsMargins(15, 10, 15, 15)
            
            descricao = QLabel(dados["descricao"])
            descricao.setStyleSheet("color: #586069; font-style: italic; padding: 5px 0;")
            main_layout.addWidget(descricao)
            
            grid = QGridLayout()
            grid.setVerticalSpacing(8)
            grid.setHorizontalSpacing(15)

            label_area = QLabel("📐 Área suprimida (ha):")
            label_area.setStyleSheet("font-weight: 500; color: #24292e;")
            area_input = QLineEdit()
            area_input.setPlaceholderText("Digite a área em hectares (ex: 5,02 ou 5.02)")
            grid.addWidget(label_area, 0, 0)
            grid.addWidget(area_input, 0, 1)

            label_auto = QLabel("📄 N° do Auto de infração:")
            label_auto.setStyleSheet("font-weight: 500; color: #24292e;")
            auto_input = QLineEdit()
            auto_input.setPlaceholderText("Preencha o número do auto de infração gerado no INÃ Fiscalização")
            grid.addWidget(label_auto, 1, 0)
            grid.addWidget(auto_input, 1, 1)

            # Só adiciona campo de embargo se não for dano
            if chave != "dano":
                label_embargo = QLabel("🚫 N° do Embargo:")
                label_embargo.setStyleSheet("font-weight: 500; color: #24292e;")
                embargo_input = QLineEdit()
                embargo_input.setPlaceholderText("Preencha o número do embargo gerado no INÃ Fiscalização")
                grid.addWidget(label_embargo, 2, 0)
                grid.addWidget(embargo_input, 2, 1)
            else:
                embargo_input = QLineEdit()
                embargo_input.setVisible(False)
            
            main_layout.addLayout(grid)
            
            if chave == "app":
                separator = QFrame()
                separator.setFrameShape(QFrame.HLine)
                separator.setStyleSheet("background-color: #e1e4e8; max-height: 1px;")
                main_layout.addWidget(separator)
                
                quadro1_box = QFrame()
                quadro1_box.setStyleSheet("""
                    QFrame {
                        background-color: #f8f9fa;
                        border-radius: 6px;
                        padding: 10px;
                    }
                """)
                quadro1_layout = QVBoxLayout()
                
                quadro1_title = QLabel("📊 QUADRO 01 - Dosimetria da Infração (ON 01/2024)")
                quadro1_title.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt; padding: 5px;")
                quadro1_layout.addWidget(quadro1_title)
                
                criterios_grid = QGridLayout()
                criterios_grid.setVerticalSpacing(10)
                
                criterios_grid.addWidget(QLabel("1. Motivo da Infração:"), 0, 0)
                self.motivo_combo = QComboBox()
                motivos = [
                    "Selecione...",
                    "Não intencional = 5",
                    "Intencional = 10",
                    "Obtenção de vantagem pecuniária = 15",
                    "Omissão deliberada em condicionantes = 20",
                    "Burla ao licenciamento ambiental = 30",
                    "Ocultamento de informações = 30",
                    "Omissão na manutenção de equipamentos = 40"
                ]
                self.motivo_combo.addItems(motivos)
                self.motivo_combo.currentTextChanged.connect(self.calcular_pontuacao_total)
                self.motivo_combo.currentTextChanged.connect(self.atualizar_faixa_a_dinamica)
                criterios_grid.addWidget(self.motivo_combo, 0, 1)
                
                criterios_grid.addWidget(QLabel("2. Consequência ambiental:"), 1, 0)
                self.consequencia_combo = QComboBox()
                consequencias = [
                    "Selecione...",
                    "Potencial = 5",
                    "Desprezível = 10",
                    "Fraca = 20",
                    "Moderada = 30",
                    "Significativa = 50",
                    "Dano não reparável = 70"
                ]
                self.consequencia_combo.addItems(consequencias)
                self.consequencia_combo.currentTextChanged.connect(self.calcular_pontuacao_total)
                self.consequencia_combo.currentTextChanged.connect(self.atualizar_faixa_a_dinamica)
                criterios_grid.addWidget(self.consequencia_combo, 1, 1)
                criterios_grid.addWidget(QLabel("3. Impacto à saúde pública:"), 2, 0)
                self.saude_combo = QComboBox()
                saude_opcoes = [
                    "Selecione...",
                    "Não houve = 0",
                    "Fraca = 5",
                    "Moderada = 10",
                    "Significativa = 20"
                ]
                self.saude_combo.addItems(saude_opcoes)
                self.saude_combo.currentTextChanged.connect(self.calcular_pontuacao_total)
                self.saude_combo.currentTextChanged.connect(self.atualizar_faixa_a_dinamica)
                criterios_grid.addWidget(self.saude_combo, 2, 1)
                quadro1_layout.addLayout(criterios_grid)
                
                pontuacao_frame = QFrame()
                pontuacao_frame.setStyleSheet("background-color: #e8f5e9; border-radius: 4px; padding: 8px;")
                pontuacao_layout = QHBoxLayout()
                pontuacao_layout.addWidget(QLabel("Pontuação Total:"))
                self.pontuacao_total_display = QLineEdit()
                self.pontuacao_total_display.setReadOnly(True)
                self.pontuacao_total_display.setStyleSheet("background-color: white; font-weight: bold; color: #006b3f;")
                self.pontuacao_total_display.setPlaceholderText("Aguardando seleção...")
                pontuacao_layout.addWidget(self.pontuacao_total_display)
                pontuacao_frame.setLayout(pontuacao_layout)
                quadro1_layout.addWidget(pontuacao_frame)
                
                quadro1_box.setLayout(quadro1_layout)
                main_layout.addWidget(quadro1_box)
                
                quadro2_box = QFrame()
                quadro2_box.setStyleSheet("""
                    QFrame {
                        background-color: #f8f9fa;
                        border-radius: 6px;
                        padding: 10px;
                    }
                """)
                quadro2_layout = QVBoxLayout()
                
                quadro2_title = QLabel("💰 QUADRO 02 - Valoração da Multa (ON 01/2024)")
                quadro2_title.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt; padding: 5px;")
                quadro2_layout.addWidget(quadro2_title)
                
                tipo_layout = QHBoxLayout()
                tipo_layout.addWidget(QLabel("Tipo de Infrator:"))
                self.tipo_infrator_combo = QComboBox()
                self.tipo_infrator_combo.addItems(["Pessoa Física", "Pessoa Jurídica"])
                self.tipo_infrator_combo.currentTextChanged.connect(self.atualizar_faixas_receita)
                tipo_layout.addWidget(self.tipo_infrator_combo)
                quadro2_layout.addLayout(tipo_layout)
                
                faixa_layout = QHBoxLayout()
                faixa_layout.addWidget(QLabel("Faixa de Receita Bruta Mensal:"))                
                self.faixa_receita_combo = QComboBox()
                self.faixas_pf = [
                    "Selecione a faixa...",
                    "Faixa A - Até 1 salário mínimo",
                    "Faixa B - 1 a 3 salários mínimos",
                    "Faixa C - 3 a 10 salários mínimos",
                    "Faixa D - 10 a 30 salários mínimos",
                    "Faixa E - 30 a 45 salários mínimos",
                    "Faixa F - Acima de 45 salários mínimos"
                ]
                self.faixas_pj = [
                    "Selecione a faixa...",
                    "Faixa A - MEI até R$ 81.000,00",
                    "Faixa B - ME até R$ 360.000,00",
                    "Faixa C - EPP até R$ 4.800.000,00",
                    "Faixa D - EMP até R$ 12.000.000,00",
                    "Faixa E - EGP-I até R$ 20.000.000,00",
                    "Faixa F - EGP-II acima de R$ 20.000.000,00"
                ]
                self.faixa_receita_combo.currentTextChanged.connect(self.atualizar_faixa_a_dinamica)
                self.faixa_receita_combo.addItems(self.faixas_pf)
                self.faixa_receita_combo.currentTextChanged.connect(
                    lambda: self.calcular_valor_multa_app(
                        "app",
                        self.areas_inputs["app"]["area"],
                        self.areas_inputs["app"]["valor"]
                    )
                )
                faixa_layout.addWidget(self.faixa_receita_combo)
                quadro2_layout.addLayout(faixa_layout)
                
                self.percentual_label = QLabel("Percentual aplicado: -")
                self.percentual_label.setStyleSheet("color: #586069; font-style: italic; padding: 5px;")
                quadro2_layout.addWidget(self.percentual_label)
                
                quadro2_box.setLayout(quadro2_layout)
                main_layout.addWidget(quadro2_box)
            
            valor_frame = QFrame()
            valor_frame.setStyleSheet("background-color: #e3f2fd; border-radius: 6px; padding: 10px;")
            valor_layout = QHBoxLayout()
            valor_layout.addWidget(QLabel("💰 Valor da Multa:"))
            valor_input = QLineEdit()
            valor_input.setReadOnly(True)
            valor_input.setStyleSheet("background-color: white; font-weight: bold; color: #006b3f; font-size: 11pt;")
            valor_input.setPlaceholderText("Clique em 'Calcular Multa'")
            valor_layout.addWidget(valor_input)
            valor_frame.setLayout(valor_layout)
            main_layout.addWidget(valor_frame)
            
            botoes_layout = QHBoxLayout()
            botoes_layout.setSpacing(10)
            
            btn_calcular = QPushButton("🔢 Calcular Multa")
            btn_detalhes = QPushButton("ℹ️ Ver detalhes")
            btn_detalhes.setObjectName("secondary")
            
            area_ref = area_input
            valor_ref = valor_input
            tipo_ref = chave
            auto_ref = auto_input
            texto_auto_ref = dados["texto_auto"]
            texto_embargo_ref = dados.get("texto_embargo", "")
            embargo_ref = embargo_input if chave != "dano" else None

            btn_calcular.clicked.connect(
                lambda checked=False, t=tipo_ref, a=area_ref, v=valor_ref:
                self.calcular_valor_multa_app(t, a, v)
                if t == "app"
                else self.calcular_multa_simples(t, a, v)
            )
            
            if chave == "app":
                cb_area = self.check_app
            elif chave == "rl":
                cb_area = self.check_rl
            elif chave == "fora":
                cb_area = self.check_fora
            elif chave == "dano":
                cb_area = self.check_dano
            
            
            btn_detalhes.clicked.connect(
                lambda checked=False, t=tipo_ref:
                self.abrir_detalhes_area(t)
            )
            
            botoes_layout.addWidget(btn_calcular)
            botoes_layout.addWidget(btn_detalhes)
            botoes_layout.addStretch()
            
            main_layout.addLayout(botoes_layout)
            
            grupo.setLayout(main_layout)
            layout.addWidget(grupo)

            mapa_checkbox = {
                "app": self.check_app,
                "rl": self.check_rl,
                "fora": self.check_fora,
                "dano": self.check_dano,
            }
            cb_area = mapa_checkbox.get(chave)
            if cb_area is not None:
                grupo.setVisible(cb_area.isChecked())
                cb_area.toggled.connect(grupo.setVisible)
            
            self.areas_inputs[chave] = {
                "area": area_input,
                "auto": auto_input,
                "embargo": embargo_input if chave != "dano" else None,
                "valor": valor_input
            }
        
        layout.addStretch()
        
        container.setLayout(layout)
        scroll.setWidget(container)
        
        final_layout = QVBoxLayout()
        final_layout.setContentsMargins(0, 0, 0, 0)
        final_layout.addWidget(scroll)
        widget.setLayout(final_layout)
        
        return widget

    def preencher_textos_auto_embargo(self, state, tipo, area_input, auto_input, embargo_input, valor_input, texto_auto_template, texto_embargo_template):
        """Preenche os campos de Auto e Embargo com os textos padrão quando o checkbox é marcado"""
        if state == Qt.Checked:  # Checkbox marcado
            # Obter os valores atuais
            area_texto = area_input.text().strip()
            valor_texto = valor_input.text().strip()
            
            # Obter dados do imóvel dos campos principais
            imovel = self.inputs["imovel"].text().strip() if "imovel" in self.inputs else ""
            municipio = self.inputs["municipio"].text().strip() if "municipio" in self.inputs else ""
            
            # Se não tiver imóvel ou município, usar valores padrão
            if not imovel:
                imovel = "NÃO INFORMADO"
            if not municipio:
                municipio = "NÃO INFORMADO"
            
            uf = self.inputs["uf"].text().strip() if "uf" in self.inputs and self.inputs["uf"].text() else "GO"
            
            # Se a área não foi preenchida, usar um placeholder
            if not area_texto:
                area_texto = "[Área a ser preenchida]"
            
            # Se o valor da multa não foi calculado, usar placeholder
            if not valor_texto:
                valor_texto = "[Valor a ser calculado]"
            
            # Substituir placeholders no texto do Auto
            texto_auto_formatado = texto_auto_template.format(
                area=area_texto,
                imovel=imovel,
                municipio=municipio,
                uf=uf,
                valor_multa=valor_texto
            )
            
            # Preencher o campo de Auto
            auto_input.setText(texto_auto_formatado)
            
            # Preencher o campo de Embargo (se existir e não for dano)
            if embargo_input and texto_embargo_template:
                texto_embargo_formatado = texto_embargo_template.format(
                    area=area_texto,
                    imovel=imovel,
                    municipio=municipio,
                    uf=uf
                )
                embargo_input.setText(texto_embargo_formatado)
            
        else:  # Checkbox desmarcado
            # Limpar os campos
            auto_input.clear()
            if embargo_input:
                embargo_input.clear()
    
    def get_percentual_pj(self, faixa, nivel):
        tabela = {
            "Faixa A": {"Nível A": (0.08, 1.0), "Nível B": (0.05, 1.0), "Nível C": (0.08, 2.0),
                       "Nível D": (0.10, 3.0), "Nível E": (0.20, 4.0)},
            "Faixa B": {"Nível A": (1.0, 2.5), "Nível B": (0.1, 2.0), "Nível C": (0.2, 3.0),
                       "Nível D": (0.4, 4.0), "Nível E": (0.8, 5.0)},
            "Faixa C": {"Nível A": (1.0, 2.5), "Nível B": (1.2, 3.0), "Nível C": (1.5, 4.0),
                       "Nível D": (3.0, 6.0), "Nível E": (5.0, 9.0)},
            "Faixa D": {"Nível A": (2.0, 4.0), "Nível B": (2.5, 5.0), "Nível C": (3.5, 6.0),
                       "Nível D": (5.5, 12.0), "Nível E": (10.0, 25.0)},
            "Faixa E": {"Nível A": (4.0, 10.0), "Nível B": (5.0, 12.0), "Nível C": (7.0, 13.0),
                       "Nível D": (10.0, 20.0), "Nível E": (20.0, 50.0)},
            "Faixa F": {"Nível A": (10.0, 30.0), "Nível B": (15.0, 40.0), "Nível C": (20.0, 50.0),
                       "Nível D": (30.0, 60.0), "Nível E": (70.0, 100.0)}
        }
        nivel_key = ""
        if "Nível A" in nivel: nivel_key = "Nível A"
        elif "Nível B" in nivel: nivel_key = "Nível B"
        elif "Nível C" in nivel: nivel_key = "Nível C"
        elif "Nível D" in nivel: nivel_key = "Nível D"
        elif "Nível E" in nivel: nivel_key = "Nível E"
        return tabela.get(faixa, {}).get(nivel_key, (0, 0))
    
    def get_percentual_por_faixa_e_nivel(self, faixa, nivel):
        percentuais = {
            "Faixa A": {"Nível A": (0.02, 0.8), "Nível B": (0.025, 0.8), "Nível C": (0.03, 1.0),
                    "Nível D": (0.035, 1.5), "Nível E": (0.045, 2.0)},
            "Faixa B": {"Nível A": (0.03, 2.0), "Nível B": (0.04, 2.5), "Nível C": (0.06, 5.0),
                    "Nível D": (0.08, 8.0), "Nível E": (0.1, 10.0)},
            "Faixa C": {"Nível A": (0.07, 2.5), "Nível B": (0.09, 3.0), "Nível C": (0.15, 4.0),
                    "Nível D": (0.3, 6.0), "Nível E": (0.4, 8.0)},
            "Faixa D": {"Nível A": (0.6, 3.0), "Nível B": (0.8, 4.0), "Nível C": (1.0, 5.0),
                    "Nível D": (1.5, 15.0), "Nível E": (3.0, 25.0)},
            "Faixa E": {"Nível A": (2.0, 6.0), "Nível B": (3.0, 8.0), "Nível C": (4.0, 10.0),
                    "Nível D": (6.0, 12.0), "Nível E": (9.0, 50.0)},
            "Faixa F": {"Nível A": (4.0, 30.0), "Nível B": (5.0, 40.0), "Nível C": (6.0, 50.0),
                    "Nível D": (8.0, 60.0), "Nível E": (12.0, 100.0)}
        }
        
        nivel_key = ""
        if "Nível A" in nivel: nivel_key = "Nível A"
        elif "Nível B" in nivel: nivel_key = "Nível B"
        elif "Nível C" in nivel: nivel_key = "Nível C"
        elif "Nível D" in nivel: nivel_key = "Nível D"
        elif "Nível E" in nivel: nivel_key = "Nível E"
        
        if faixa in percentuais and nivel_key in percentuais[faixa]:
            return percentuais[faixa][nivel_key]
        return (0, 0)

    def calcular_pontuacao_total(self):
        try:
            if not hasattr(self, 'motivo_combo'):
                return 0, ""
            
            motivo_pontos = 0
            consequencia_pontos = 0
            saude_pontos = 0
            
            motivo_text = self.motivo_combo.currentText()
            if "=" in motivo_text and motivo_text != "Selecione...":
                try:
                    motivo_pontos = int(motivo_text.split("=")[1].strip())
                except:
                    motivo_pontos = 0
            
            consequencia_text = self.consequencia_combo.currentText()
            if "=" in consequencia_text and consequencia_text != "Selecione...":
                try:
                    consequencia_pontos = int(consequencia_text.split("=")[1].strip())
                except:
                    consequencia_pontos = 0
            
            saude_text = self.saude_combo.currentText()
            if "=" in saude_text and saude_text != "Selecione...":
                try:
                    saude_pontos = int(saude_text.split("=")[1].strip())
                except:
                    saude_pontos = 0
            
            total = motivo_pontos + consequencia_pontos + saude_pontos
            
            if 10 <= total <= 20:
                nivel = "Nível A (Gravidade Baixa)"
            elif 21 <= total <= 40:
                nivel = "Nível B (Gravidade Média-Baixa)"
            elif 41 <= total <= 60:
                nivel = "Nível C (Gravidade Média)"
            elif 61 <= total <= 80:
                nivel = "Nível D (Gravidade Média-Alta)"
            elif 81 <= total <= 100:
                nivel = "Nível E (Gravidade Alta)"
            else:
                nivel = ""
            
            if total >= 10:
                self.pontuacao_total_display.setText(f"{total} pontos - {nivel}")
                self.pontuacao_total_display.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold;")
                return total, nivel
            else:
                if total > 0:
                    self.pontuacao_total_display.setText(f"{total} pontos - Pontuação insuficiente (mínimo 10)")
                    self.pontuacao_total_display.setStyleSheet("background-color: #fff3cd; color: #856404; font-weight: bold;")
                else:
                    self.pontuacao_total_display.setText("")
                    self.pontuacao_total_display.setPlaceholderText("Aguardando seleção...")
                    self.pontuacao_total_display.setStyleSheet("background-color: white;")
                return total, ""
                
        except Exception as e:
            print(f"Erro ao calcular pontuação: {e}")
            return 0, ""

    def obter_parametros_valoracao_app(self):
        """
        Lê o que o usuário efetivamente selecionou nos combos da dosimetria
        (Quadro 01) e da valoração (Quadro 02) da área de APP e devolve os
        parâmetros já calculados (nível, faixa, percentuais).
        """
        pontuacao, nivel_completo = self.calcular_pontuacao_total()
        nivel_curto = nivel_completo.split(" (")[0] if " (" in nivel_completo else (nivel_completo or "Nível B")

        tipo_infrator = self.tipo_infrator_combo.currentText() if hasattr(self,
                                                                          'tipo_infrator_combo') else "Pessoa Física"

        faixa_texto = self.faixa_receita_combo.currentText() if hasattr(self, 'faixa_receita_combo') else ""
        if faixa_texto and "Selecione" not in faixa_texto:
            faixa = faixa_texto.split(" - ")[0]
            descricao_faixa = faixa_texto.split(" - ")[1] if " - " in faixa_texto else ""
        else:
            faixa = "Faixa A"
            descricao_faixa = "Até 1 salário mínimo" if tipo_infrator == "Pessoa Física" else "MEI até R$ 81.000,00"

        if tipo_infrator == "Pessoa Jurídica":
            perc_min, perc_max = self.get_percentual_pj(faixa, nivel_curto)
        else:
            perc_min, perc_max = self.get_percentual_por_faixa_e_nivel(faixa, nivel_curto)

        return {
            "pontuacao": pontuacao,
            "nivel": nivel_curto,
            "tipo_infrator": tipo_infrator,
            "faixa": faixa,
            "descricao_faixa": descricao_faixa,
            "perc_min": perc_min,
            "perc_max": perc_max,
        }

    def atualizar_faixa_a_dinamica(self):
        """
        Atualiza a mensagem da Faixa A dinamicamente baseada na pontuação atual
        e exibe o percentual aplicado conforme a combinação Faixa + Nível
        """
        try:
            # Verifica se a faixa selecionada é a Faixa A
            if not hasattr(self, 'faixa_receita_combo'):
                return
                
            faixa_texto = self.faixa_receita_combo.currentText()
            if "Faixa A" not in faixa_texto or "Selecione" in faixa_texto:
                return
            
            # Calcula a pontuação atual e obtém o Nível
            pontuacao, nivel = self.calcular_pontuacao_total()
            
            if pontuacao < 10:
                if hasattr(self, 'percentual_label'):
                    self.percentual_label.setText(
                        "⚠️ Aguardando pontuação mínima de 10 pontos para definir o Nível."
                    )
                return
            
            # Extrai o Nível (ex: "Nível B (Gravidade Média-Baixa)" -> "Nível B")
            nivel_curto = nivel.split(" (")[0] if " (" in nivel else nivel
            
            # Obtém os percentuais para Faixa A + Nível atual
            tipo_infrator = self.tipo_infrator_combo.currentText()
            
            if tipo_infrator == "Pessoa Jurídica":
                perc_min, perc_max = self.get_percentual_pj("Faixa A", nivel_curto)
            else:
                perc_min, perc_max = self.get_percentual_por_faixa_e_nivel("Faixa A", nivel_curto)
            
            # Verifica se UC está marcado para adicionar informação
            uc_texto = " (DOBRO - Art. 93 - UC)" if self.is_uc_marcado() else ""
            
            # Monta a mensagem completa
            mensagem = (
                f"📊 PERCENTUAL APLICADO: {perc_min:.3f}% a {perc_max:.3f}%{uc_texto}\n\n"
                f"⚠️ Não foi possível identificar a capacidade econômica do autuado pela ausência "
                f"de documentos ou informações. Neste caso, optou-se pela FAIXA A (Receita Bruta "
                f"mensal de até 1 salário mínimo, {nivel_curto}, mínimo + {perc_min:.3f}% até {perc_max:.3f}% do teto). "
                f"No entanto, no momento da audiência de autocomposição poderá ser reclassificado "
                f"conforme a capacidade econômica, mediante comprovação por documentos, "
                f"conforme § 2º do Art. 8 da ON - 01/2024."
            )
            
            if hasattr(self, 'percentual_label'):
                self.percentual_label.setText(mensagem)
                self.percentual_label.setWordWrap(True)
                self.percentual_label.setStyleSheet(
                    "color: #856404; background-color: #fff3cd; "
                    "border: 1px solid #ffeeba; border-radius: 6px; "
                    "padding: 10px; font-style: normal; font-size: 9pt;"
                )
            
            # Se houver área preenchida, recalcula automaticamente a multa
            if "app" in self.areas_inputs:
                area_input = self.areas_inputs["app"].get("area")
                valor_input = self.areas_inputs["app"].get("valor")
                if area_input and area_input.text().strip():
                    self.calcular_valor_multa_app("app", area_input, valor_input)
                    
        except Exception as e:
            print(f"Erro ao atualizar Faixa A dinâmica: {e}")

    def gerar_novos_relatorios(self):
        """Gera os novos relatórios selecionados com dados integrados de todas as abas"""
        selecionados = []
        if self.check_relatorio_ina and self.check_relatorio_ina.isChecked():
            selecionados.append(("relatorio_ina", "Relatório Inã - Supressão de vegetação nativa"))
        if self.check_despacho and self.check_despacho.isChecked():
            selecionados.append(("despacho_autocomposicao", "Despacho Autocomposição"))
        if self.check_minuta and self.check_minuta.isChecked():
            selecionados.append(("modelo_minuta", "Modelo de Minuta"))
        if self.check_autodenuncia and self.check_autodenuncia.isChecked():
            selecionados.append(("relatorio_autodenuncia", "Relatório Autodenúncia"))

        # NOVOS RELATÓRIOS
        if hasattr(self, 'check_barramento') and self.check_barramento.isChecked():
            selecionados.append(("barramento", "Relatório INÃ - Barramento"))
        if hasattr(self, 'check_parcelamento') and self.check_parcelamento.isChecked():
            selecionados.append(("parcelamento", "Relatório INÃ - Parcelamento"))

        if not selecionados:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um relatório para gerar!")
            return

        if not self.inputs["processo"].text() or not self.inputs["imovel"].text():
            resposta = QMessageBox.question(self, "Aviso",
                                            "Processo e Imóvel são obrigatórios. Continuar mesmo assim?",
                                            QMessageBox.Yes | QMessageBox.No)
            if resposta == QMessageBox.No:
                return

        now = datetime.now()

        # ==========================================================
        # COLETA DE DADOS - ABA DADOS GERAIS
        # ==========================================================
        app_area_texto = self.areas_inputs.get("app", {}).get("area", QLineEdit()).text() or "0"
        rl_area_texto = self.areas_inputs.get("rl", {}).get("area", QLineEdit()).text() or "0"
        fora_area_texto = self.areas_inputs.get("fora", {}).get("area", QLineEdit()).text() or "0"
        dano_area_texto = self.areas_inputs.get("dano", {}).get("area", QLineEdit()).text() or "0"

        app_valor_texto = self.areas_inputs.get("app", {}).get("valor", QLineEdit()).text() or "NÃO CALCULADO"
        rl_valor_texto = self.areas_inputs.get("rl", {}).get("valor", QLineEdit()).text() or "NÃO CALCULADO"
        fora_valor_texto = self.areas_inputs.get("fora", {}).get("valor", QLineEdit()).text() or "NÃO CALCULADO"
        dano_valor_texto = self.areas_inputs.get("dano", {}).get("valor", QLineEdit()).text() or "NÃO CALCULADO"

        # ==========================================================
        # COLETA DE DADOS - ABA DADOS PARA O RELATÓRIO (BARRAMENTO)
        # ==========================================================
        dados_barramento = {}
        if hasattr(self, 'campos_barramento'):
            for nome, campo in self.campos_barramento.items():
                if hasattr(campo, 'toPlainText'):
                    dados_barramento[nome] = campo.toPlainText()
                elif hasattr(campo, 'currentText'):
                    dados_barramento[nome] = campo.currentText()
                else:
                    dados_barramento[nome] = campo.text()

        # ==========================================================
        # COLETA DE DADOS - ABA DADOS PARA O RELATÓRIO (PARCELAMENTO)
        # ==========================================================
        dados_parcelamento = {}
        if hasattr(self, 'campos_parcelamento'):
            for nome, campo in self.campos_parcelamento.items():
                if hasattr(campo, 'toPlainText'):
                    dados_parcelamento[nome] = campo.toPlainText()
                elif hasattr(campo, 'currentText'):
                    dados_parcelamento[nome] = campo.currentText()
                else:
                    dados_parcelamento[nome] = campo.text()

        # ==========================================================
        # DICIONÁRIO PRINCIPAL DE DADOS
        # ==========================================================
        dados = {
            "processo": self.inputs["processo"].text() or "NÃO INFORMADO",
            "data": now.strftime("%d/%m/%Y"),
            "imovel": self.inputs["imovel"].text() or "NÃO INFORMADO",
            "municipio": self.inputs["municipio"].text() or "NÃO INFORMADO",
            "uf": self.inputs["uf"].text() or "GO",
            "car": self.inputs["car"].text() or "NÃO INFORMADO",
            "proprietario": self.inputs["proprietario"].text() or "NÃO INFORMADO",
            "cpf": self.inputs["cpf"].text() or "NÃO INFORMADO",
            "coordenadas": self.inputs["coordenadas"].text() or "NÃO INFORMADO",
            "os": self.inputs.get("os", QLineEdit()).text() or "NÃO INFORMADO",
            "observacoes": self.obs.toPlainText(),
            "alertas": self.inputs.get("alertas", QLineEdit()).text() or "NÃO INFORMADO",
            "intervalo_supressao": self.inputs.get("intervalo_supressao",
                                                   QLineEdit()).text() or "período não informado",

            # Dados das áreas
            "app_area": app_area_texto,
            "rl_area": rl_area_texto,
            "fora_area": fora_area_texto,
            "dano_area": dano_area_texto,

            "valor_app": app_valor_texto,
            "valor_rl": rl_valor_texto,
            "valor_fora": fora_valor_texto,
            "valor_dano": dano_valor_texto,

            "app_auto": self.areas_inputs.get("app", {}).get("auto", QLineEdit()).text() or "",
            "rl_auto": self.areas_inputs.get("rl", {}).get("auto", QLineEdit()).text() or "",
            "fora_auto": self.areas_inputs.get("fora", {}).get("auto", QLineEdit()).text() or "",
            "dano_auto": self.areas_inputs.get("dano", {}).get("auto", QLineEdit()).text() or "",

            "app_embargo": self.areas_inputs.get("app", {}).get("embargo", QLineEdit()).text() or "",
            "rl_embargo": self.areas_inputs.get("rl", {}).get("embargo", QLineEdit()).text() or "",
            "fora_embargo": self.areas_inputs.get("fora", {}).get("embargo", QLineEdit()).text() or "",

            # Dados do relatório INÃ (já existentes)
            "processo_sei": self.inputs.get("processo", QLineEdit()).text() or "",
            "processo_ina": self.inputs.get("processo", QLineEdit()).text() or "",

            # ==========================================================
            # DADOS DO BARRAMENTO (coletados da aba específica)
            # ==========================================================
            "data_ocorrencia": dados_barramento.get("data_ocorrencia", now.strftime("%d/%m/%Y")),
            "manifestacao": dados_barramento.get("manifestacao", "NÃO INFORMADO"),
            "fato_denunciado": dados_barramento.get("fato_denunciado", self.obs.toPlainText() or "NÃO INFORMADO"),
            "texto_ina_pr": dados_barramento.get("texto_ina_pr", "NÃO INFORMADO"),
            "data_sobrevoo": dados_barramento.get("data_sobrevoo", "NÃO INFORMADO"),
            "mapa_geral": dados_barramento.get("mapa_geral", "NÃO INFORMADO"),
            "mapa_temporal": dados_barramento.get("mapa_temporal", "NÃO INFORMADO"),
            "area_barramento": dados_barramento.get("area_barramento", "0"),
            "area_supressao": dados_barramento.get("area_supressao", "0"),
            "area_supressao_arredondada": dados_barramento.get("area_supressao_arredondada", "0"),
            "intervalo_supressao_barramento": dados_barramento.get("intervalo_supressao", "período não informado"),
            "auto_barramento": dados_barramento.get("auto_barramento", ""),
            "embargo_barramento": dados_barramento.get("embargo_barramento", ""),
            "auto_artigo_66": dados_barramento.get("auto_artigo_66", ""),
            "valor_ai_66": float(dados_barramento.get("valor_ai_66", "3000").replace(',', '.') or 3000),
            "auto_artigo_43": dados_barramento.get("auto_artigo_43", ""),
            "valor_ai_43": float(dados_barramento.get("valor_ai_43", "5012.50").replace(',', '.') or 5012.50),
            "valor_recursos_hidricos": float(
                dados_barramento.get("valor_recursos_hidricos", "902.25").replace(',', '.') or 902.25),

            # ==========================================================
            # DADOS DO PARCELAMENTO (coletados da aba específica)
            # ==========================================================
            "parc_data_ocorrencia": dados_parcelamento.get("parc_data_ocorrencia", now.strftime("%d/%m/%Y")),
            "parc_manifestacao": dados_parcelamento.get("parc_manifestacao", "NÃO INFORMADO"),
            "parc_area": dados_parcelamento.get("parc_area", "0"),
            "parc_auto": dados_parcelamento.get("parc_auto", ""),
            "parc_embargo": dados_parcelamento.get("parc_embargo", ""),
            "parc_valor_multa": float(dados_parcelamento.get("parc_valor_multa", "3500").replace(',', '.') or 3500),
            "parc_tipo_infrator": dados_parcelamento.get("parc_tipo_infrator", "Pessoa Física"),
            "parc_faixa_receita": dados_parcelamento.get("parc_faixa_receita", "Faixa A"),
            "parc_nivel_gravidade": dados_parcelamento.get("parc_nivel_gravidade", "Nível C"),
            "parc_percentual": dados_parcelamento.get("parc_percentual", "0,03%"),
            "parc_qtd_lotes": dados_parcelamento.get("parc_qtd_lotes", "NÃO INFORMADO"),
            "parc_app_info": dados_parcelamento.get("parc_app_info", "NÃO INFORMADO"),
            "parc_fato_denunciado": dados_parcelamento.get("parc_fato_denunciado", "NÃO INFORMADO"),

        }

        # ==========================================================
        # SELECIONA PASTA PARA SALVAR
        # ==========================================================
        pasta = QFileDialog.getExistingDirectory(self, "Escolha a pasta para salvar os relatórios")
        if not pasta:
            return

        arquivos_gerados = []

        # ==========================================================
        # GERAÇÃO DOS RELATÓRIOS
        # ==========================================================
        for tipo, nome in selecionados:
            if tipo == "relatorio_ina":
                conteudo = self.gerar_relatorio_ina(dados)
                nome_arquivo = f"Relatorio_Ina_Supressao_{dados['imovel']}_{now.strftime('%Y%m%d_%H%M%S')}.txt"
            elif tipo == "despacho_autocomposicao":
                conteudo = self.gerar_despacho_autocomposicao(dados)
                nome_arquivo = f"Despacho_Autocomposicao_{dados['imovel']}_{now.strftime('%Y%m%d_%H%M%S')}.txt"
            elif tipo == "modelo_minuta":
                conteudo = self.gerar_modelo_minuta(dados)
                nome_arquivo = f"Modelo_Minuta_{dados['imovel']}_{now.strftime('%Y%m%d_%H%M%S')}.txt"
            elif tipo == "relatorio_autodenuncia":
                conteudo = self.gerar_relatorio_autodenuncia(dados)
                nome_arquivo = f"Relatorio_Autodenuncia_{dados['imovel']}_{now.strftime('%Y%m%d_%H%M%S')}.txt"
            elif tipo == "barramento":
                # Usa o gerenciador com os dados coletados
                conteudo = self.gerenciador.gerar_relatorio_barramento(dados)
                nome_arquivo = f"Relatorio_Barramento_{dados['imovel']}_{now.strftime('%Y%m%d_%H%M%S')}.txt"
            elif tipo == "parcelamento":
                # Usa o método da própria classe com os dados coletados
                conteudo = self.gerar_relatorio_parcelamento(dados)
                nome_arquivo = f"Relatorio_Parcelamento_{dados['imovel']}_{now.strftime('%Y%m%d_%H%M%S')}.txt"
            else:
                continue

            # Remove caracteres inválidos para nome de arquivo
            nome_arquivo = "".join(c for c in nome_arquivo if c.isalnum() or c in '._- ')

            caminho_arquivo = os.path.join(pasta, nome_arquivo)

            with open(caminho_arquivo, 'w', encoding='utf-8') as f:
                f.write(conteudo)

            arquivos_gerados.append(caminho_arquivo)

        # ==========================================================
        # MENSAGEM FINAL
        # ==========================================================
        if arquivos_gerados:
            QMessageBox.information(self, "Sucesso", f"{len(arquivos_gerados)} relatório(s) gerado(s) com sucesso!")
        else:
            QMessageBox.warning(self, "Erro", "Nenhum relatório foi gerado!")

    def calcular_valor_multa_app(self, tipo, area_input, valor_input):
        try:
            if area_input is None:
                QMessageBox.warning(self, "Erro", "Campo de área não encontrado!")
                return
            
            area_original = self.validar_area(area_input)
            if area_original is None:
                return
            
            area = math.ceil(area_original)
            
            if int(area) != float(area_original):
                resposta = QMessageBox.question(self, "Arredondamento de área",
                    f"Área original: {area_original} ha\nÁrea arredondada para cima: {area} ha\n\nDeseja continuar?",
                    QMessageBox.Yes | QMessageBox.No)
                if resposta == QMessageBox.No:
                    return
                area_input.setText(str(area))
            else:
                area = area_original
            
            valor_base_ha = 5000.00
            teto_ha = 50000.00
            
            if not hasattr(self, 'motivo_combo'):
                QMessageBox.warning(self, "Erro", "Sistema não inicializado corretamente.")
                return
            
            pontuacao, nivel = self.calcular_pontuacao_total()
            
            if pontuacao < 10:
                QMessageBox.warning(self, "Aviso", "A pontuação mínima para cálculo é 10 pontos.")
                return
            
            faixa_texto = self.faixa_receita_combo.currentText()
            if "Selecione" in faixa_texto or not faixa_texto:
                QMessageBox.warning(self, "Aviso", "Selecione a faixa de receita bruta mensal.")
                return
            
            faixa = faixa_texto.split(" - ")[0]
            tipo_infrator = self.tipo_infrator_combo.currentText()
            
            # Extrai o Nível curto
            nivel_curto = nivel.split(" (")[0] if " (" in nivel else nivel

            if tipo_infrator == "Pessoa Jurídica":
                perc_min, perc_max = self.get_percentual_pj(faixa, nivel_curto)
            else:
                perc_min, perc_max = self.get_percentual_por_faixa_e_nivel(faixa, nivel_curto)
            
            acrescimo_min_ha = teto_ha * (perc_min / 100)
            acrescimo_max_ha = teto_ha * (perc_max / 100)
            
            valor_por_ha_min = valor_base_ha + acrescimo_min_ha
            valor_por_ha_max = valor_base_ha + acrescimo_max_ha
            
            valor_total_min = area * valor_por_ha_min
            valor_total_max = area * valor_por_ha_max
            
            # VERIFICA SE UC ESTÁ MARCADO - APLICA DOBRO
            if self.is_uc_marcado():
                valor_total_min *= 2
                valor_total_max *= 2
            
            valor_total_min_str = f"R$ {valor_total_min:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
            valor_total_max_str = f"R$ {valor_total_max:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
            
            # Se for Faixa A, exibe a mensagem especial com o percentual atualizado
            if faixa == "Faixa A":
                uc_texto = " (DOBRO - Art. 93 - UC)" if self.is_uc_marcado() else ""
                mensagem = (
                    f"📊 PERCENTUAL APLICADO: {perc_min:.3f}% a {perc_max:.3f}%{uc_texto}\n\n"
                    f"⚠️ Não foi possível identificar a capacidade econômica do autuado pela ausência "
                    f"de documentos ou informações. Neste caso, optou-se pela FAIXA A (Receita Bruta "
                    f"mensal de até 1 salário mínimo, {nivel_curto}, mínimo + {perc_min:.3f}% até {perc_max:.3f}% do teto). "
                    f"No entanto, no momento da audiência de autocomposição poderá ser reclassificado "
                    f"conforme a capacidade econômica, mediante comprovação por documentos, "
                    f"conforme § 2º do Art. 8 da ON - 01/2024."
                )
                if hasattr(self, 'percentual_label'):
                    self.percentual_label.setText(mensagem)
                    self.percentual_label.setWordWrap(True)
                    self.percentual_label.setStyleSheet(
                        "color: #856404; background-color: #fff3cd; "
                        "border: 1px solid #ffeeba; border-radius: 6px; "
                        "padding: 10px; font-style: normal; font-size: 9pt;"
                    )
            else:
                # Para outras faixas, exibe apenas o percentual
                if hasattr(self, 'percentual_label'):
                    uc_texto = " (DOBRO - Art. 93 - UC)" if self.is_uc_marcado() else ""
                    self.percentual_label.setText(
                        f"✓ Área: {area} ha | Nível: {nivel_curto} | Faixa: {faixa} | Percentual: {perc_min:.3f}% a {perc_max:.3f}%{uc_texto}"
                    )
                    self.percentual_label.setWordWrap(False)
                    self.percentual_label.setStyleSheet("color: #586069; font-style: italic; padding: 5px;")
            
            valor_input.setText(f"{valor_total_min_str} a {valor_total_max_str}")
            valor_input.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold; font-size: 11pt;")
            
            self.atualizar_texto_supressao_apos_calculo("app", area_input, valor_input)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro inesperado", f"Erro ao calcular: {str(e)}")

    def calcular_multa_simples(self, tipo, area_input, valor_input):
        try:
            if area_input is None:
                QMessageBox.warning(self, "Erro", "Campo de área não encontrado!")
                return
            
            texto_area = area_input.text().strip()
            if not texto_area:
                QMessageBox.warning(self, "Erro", "Por favor, digite um valor para a área!")
                return
            
            texto_area = texto_area.replace(",", ".")
            
            try:
                area_original = float(texto_area)
            except ValueError:
                QMessageBox.warning(self, "Erro de conversão", f"Valor inválido: '{texto_area}'")
                return
            
            if area_original <= 0:
                QMessageBox.warning(self, "Erro", "A área deve ser maior que zero!")
                return
            
            area = math.ceil(area_original)
            
            if area != area_original:
                resposta = QMessageBox.question(self, "Arredondamento de área",
                    f"Área original: {area_original} ha\nÁrea arredondada para cima: {area} ha\n\nDeseja continuar?",
                    QMessageBox.Yes | QMessageBox.No)
                if resposta == QMessageBox.No:
                    return
                area_input.setText(str(area))
            else:
                area = area_original
            
            valores = {"rl": 5000.00, "fora": 1000.00, "dano": 300.00}
            valor_por_ha = valores.get(tipo, 1000.00)
            valor_total = area * valor_por_ha
            
            # VERIFICA SE UC ESTÁ MARCADO - APLICA DOBRO
            if self.is_uc_marcado():
                valor_total *= 2
            
            valor_total_str = f"R$ {valor_total:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
            valor_input.setText(valor_total_str)
            
            # Adiciona indicador de dobro
            if self.is_uc_marcado():
                valor_input.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold; font-size: 11pt; border: 2px solid #ff6600;")
            else:
                valor_input.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold; font-size: 11pt;")
            
            self.atualizar_texto_supressao_apos_calculo(tipo, area_input, valor_input)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro inesperado", f"Erro ao calcular: {str(e)}")

    def atualizar_texto_supressao_apos_calculo(self, tipo, area_input, valor_input):
        """Atualiza apenas o valor da multa, sem preencher Auto e Embargo"""
        # Este método agora só atualiza o valor da multa, não mexe nos campos Auto/Embargo
        pass
    
    def obter_template_texto(self, tipo):
        templates = {
            "app": "SUPRESSÃO PASSÍVEL - APP\n\nPor suprimir {area} hectares de vegetação nativa em Área de Preservação Permanente (APP), janeiro de 2025 a abril de 2025, sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}.\nValor da Multa em Área Passível: {valor_multa}\n\nEMBARGO PASSÍVEL\nFica embargada a área em APP, por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}.",
            "rl": "SUPRESSÃO PASSÍVEL - RESERVA LEGAL\n\nPor suprimir {area} hectares de vegetação nativa em Reserva Legal (RL), janeiro de 2025 a abril de 2025, sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}.\nValor da Multa em Área Passível: {valor_multa}\n\nEMBARGO PASSÍVEL\nFica embargada a área em Reserva Legal (RL), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}.",
            "fora": "SUPRESSÃO PASSÍVEL\n\nPor suprimir {area} hectares de vegetação nativa em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), janeiro de 2025 a abril de 2025, sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}.\nValor da Multa em Área Passível: {valor_multa}\n\nEMBARGO PASSÍVEL\nFica embargada a área em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP), por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}.",
            "dano": "SUPRESSÃO PASSÍVEL - DANO AMBIENTAL\n\nPor causar dano ambiental em {area} hectares de vegetação nativa, janeiro de 2025 a abril de 2025, sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}.\nValor da Multa em Área Passível: {valor_multa}\n\nEMBARGO PASSÍVEL\nFica embargada a área com dano ambiental, por supressão de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."
        }
        return templates.get(tipo, templates["fora"])

    def abrir_detalhes_area(self, tipo):
        conteudos = {
            "app": "APP - Área de Preservação Permanente\n\nÁreas protegidas por lei ambiental.\n\nMulta: R$ 5.000 a R$ 50.000 por hectare.",
            "rl": "Reserva Legal\n\nÁrea destinada à conservação da vegetação nativa.\n\nMulta: R$ 5.000 por hectare.",
            "fora": "Fora da Reserva Legal\n\nSupressão fora da área de RL.\n\nMulta: R$ 1.000 por hectare.",
            "dano": "Dano Ambiental\n\nDegradação ambiental identificada durante fiscalização.\n\nMulta: R$ 300 por hectare."
        }
        titulos = {"app": "APP", "rl": "Reserva Legal", "fora": "Fora da RL", "dano": "Dano Ambiental"}
        modal = ModalDetalhes(titulos[tipo], conteudos[tipo], self)
        modal.exec()

    def criar_aba_areas_selecionadas(self):
        """Cria a aba com a pré-visualização das áreas selecionadas"""
        widget = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)

        # Seção de pré-visualização
        grupo_preview = QGroupBox("PRÉ-VISUALIZAÇÃO DAS ÁREAS SELECIONADAS")
        grupo_preview.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 2px solid #006b3f;
                border-radius: 8px;
                margin-top: 15px;
            }
            QGroupBox::title {
                color: #006b3f;
                font-weight: bold;
                font-size: 12pt;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
            }
        """)

        grupo_preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        preview_layout = QVBoxLayout()

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMinimumHeight(500)
        self.preview_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_text.setPlaceholderText(
            "Selecione áreas na aba 'Autos e Embargos' para visualizar a pré-visualização...")

        self.preview_text.setStyleSheet("""
            QTextEdit {
                font-family: 'Courier New', 'Consolas', monospace;
                font-size: 11pt;
                line-height: 1.6;
                background-color: #fafafa;
            }
        """)

        preview_layout.addWidget(self.preview_text)

        # Layout dos botões
        botoes_layout = QHBoxLayout()
        botoes_layout.setSpacing(10)

        btn_atualizar_preview = QPushButton("🔄 ATUALIZAR PRÉ-VISUALIZAÇÃO")
        btn_atualizar_preview.setStyleSheet("""
            QPushButton {
                background-color: #006b3f;
                font-size: 12pt;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #008c52;
            }
        """)
        btn_atualizar_preview.clicked.connect(self.atualizar_preview)

        # NOVO BOTÃO - GERAR TXT
        btn_gerar_txt = QPushButton("📄 GERAR TXT DO PREVIEW")
        btn_gerar_txt.setStyleSheet("""
            QPushButton {
                background-color: #ff6b00;
                font-size: 12pt;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e85a00;
            }
        """)
        btn_gerar_txt.clicked.connect(self.gerar_txt_preview)

        botoes_layout.addWidget(btn_atualizar_preview)
        botoes_layout.addWidget(btn_gerar_txt)
        botoes_layout.addStretch()

        preview_layout.addLayout(botoes_layout)

        info_label = QLabel(
            "ℹ️ As informações exibidas aqui são geradas a partir dos dados preenchidos na aba 'Autos e Embargos'")
        info_label.setStyleSheet("color: #586069; font-style: italic; padding: 5px;")
        info_label.setAlignment(Qt.AlignCenter)
        preview_layout.addWidget(info_label)

        grupo_preview.setLayout(preview_layout)
        layout.addWidget(grupo_preview)

        container.setLayout(layout)
        scroll.setWidget(container)

        final_layout = QVBoxLayout()
        final_layout.addWidget(scroll)
        widget.setLayout(final_layout)

        return widget

    def criar_aba_artigos(self):
        """Cria a aba com os artigos (48, 66, 79 e um artigo genérico)"""
        widget = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        self.artigos_inputs = {}

        # ==========================================================
        # CHECKBOX UC - UNIDADE DE CONSERVAÇÃO (mesma opção da aba
        # "Autos e Embargos (Supressão de vegetação nativa)", sincronizada)
        # ==========================================================
        grupo_uc_artigos = QGroupBox("📍 Unidade de Conservação")
        grupo_uc_artigos.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 1px solid #e1e4e8;
                border-radius: 8px;
            }
        """)
        uc_artigos_layout = QHBoxLayout()
        uc_artigos_layout.setSpacing(20)
        uc_artigos_layout.setContentsMargins(15, 10, 15, 10)

        self.check_uc_artigos = QCheckBox("UC - Unidade de Conservação")
        self.check_uc_artigos.setStyleSheet("""
            QCheckBox {
                spacing: 8px;
                font-weight: 500;
            }
        """)
        self.check_uc_artigos.stateChanged.connect(
            lambda state: self.sincronizar_checkbox_uc(self.check_uc_artigos, state))

        uc_artigos_layout.addWidget(self.check_uc_artigos)
        uc_artigos_layout.addStretch()
        grupo_uc_artigos.setLayout(uc_artigos_layout)
        layout.addWidget(grupo_uc_artigos)

        # ==========================================================
        # CONFIGURAÇÃO DOS ARTIGOS FIXOS (48, 66, 79)
        # ==========================================================
        artigos_config = {
            "art48": {
                "titulo": "Art. 48 - Impedir regeneração de vegetação",
                "descricao": "Multa fechada: R$ 5.000,00 por hectare (Decreto 6514/2008)",
                "tipo_calculo": "simples",
                "valor_por_ha": 5000.00,
                "mostra_area": True
            },
            "art66": {
                "titulo": "Art. 66 - Executar atividade sem licença",
                "descricao": "Multa aberta: Valoração por dosimetria (ON 01/2024)",
                "tipo_calculo": "dosimetria",
                "valor_base": 500.00,
                "teto": 10000000.00,
                "mostra_area": False
            },
            "art79": {
                "titulo": "Art. 79 - Descumprimento de embargo",
                "descricao": "Multa aberta: Valoração por dosimetria (ON 01/2024)",
                "tipo_calculo": "dosimetria",
                "valor_base": 10000.00,
                "teto": 10000000.00,
                "mostra_area": False
            }
        }

        for chave, dados in artigos_config.items():
            grupo = QGroupBox(dados["titulo"])
            grupo.setStyleSheet(f"""
                QGroupBox {{
                    background-color: #f8f9fa;
                    border: 1px solid #e1e4e8;
                    border-radius: 8px;
                    margin-top: 10px;
                    padding-top: 15px;
                }}
                QGroupBox::title {{
                    font-weight: 600;
                    color: #006b3f;
                }}
            """)

            main_layout = QVBoxLayout()
            main_layout.setSpacing(10)
            main_layout.setContentsMargins(15, 10, 15, 15)

            descricao = QLabel(dados["descricao"])
            descricao.setStyleSheet("color: #586069; font-style: italic; padding: 5px 0;")
            main_layout.addWidget(descricao)

            grid = QGridLayout()
            grid.setVerticalSpacing(8)
            grid.setHorizontalSpacing(15)

            # ==========================================================
            # CAMPO ÁREA - APENAS PARA ART. 48
            # ==========================================================
            area_input = QLineEdit()

            if dados.get("mostra_area", True):
                label_area = QLabel("📐 Área suprimida (ha):")
                label_area.setStyleSheet("font-weight: 500; color: #24292e;")
                area_input.setPlaceholderText("Digite a área em hectares (ex: 5,02 ou 5.02)")
                area_input.setMinimumWidth(200)
                grid.addWidget(label_area, 0, 0)
                grid.addWidget(area_input, 0, 1)
            else:
                area_input.setVisible(False)
                info_label = QLabel("ℹ️ Este artigo não depende de área para cálculo da multa")
                info_label.setStyleSheet("color: #006b3f; font-style: italic; padding: 5px 0;")
                main_layout.addWidget(info_label)

            # -------- AUTO --------
            label_auto = QLabel("📄 Nº Auto de Infração:")
            label_auto.setStyleSheet("font-weight: 500; color: #24292e;")
            auto_input = QLineEdit()
            auto_input.setPlaceholderText("Número do auto")
            grid.addWidget(label_auto, 1, 0)
            grid.addWidget(auto_input, 1, 1)

            # -------- EMBARGO --------
            label_embargo = QLabel("🚫 Nº Embargo:")
            label_embargo.setStyleSheet("font-weight: 500; color: #24292e;")
            embargo_input = QLineEdit()
            embargo_input.setPlaceholderText("Número do embargo")
            grid.addWidget(label_embargo, 2, 0)
            grid.addWidget(embargo_input, 2, 1)

            main_layout.addLayout(grid)

            # ==========================================================
            # SE FOR DOSIMETRIA (ART. 66 E 79)
            # ==========================================================
            if dados["tipo_calculo"] == "dosimetria":
                separator = QFrame()
                separator.setFrameShape(QFrame.HLine)
                separator.setStyleSheet("background-color: #e1e4e8; max-height: 1px;")
                main_layout.addWidget(separator)

                # -------- QUADRO 01 - DOSIMETRIA --------
                quadro1_box = QFrame()
                quadro1_box.setStyleSheet("""
                    QFrame {
                        background-color: #ffffff;
                        border-radius: 6px;
                        padding: 10px;
                        border: 1px solid #e1e4e8;
                    }
                """)
                quadro1_layout = QVBoxLayout()

                quadro1_title = QLabel("📊 QUADRO 01 - Dosimetria da Infração (ON 01/2024)")
                quadro1_title.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt; padding: 5px;")
                quadro1_layout.addWidget(quadro1_title)

                criterios_grid = QGridLayout()
                criterios_grid.setVerticalSpacing(10)

                # Motivo
                criterios_grid.addWidget(QLabel("1. Motivo da Infração:"), 0, 0)
                motivo_combo = QComboBox()
                motivos = [
                    "Selecione...",
                    "Não intencional = 5",
                    "Intencional = 10",
                    "Obtenção de vantagem pecuniária = 15",
                    "Omissão deliberada em condicionantes = 20",
                    "Burla ao licenciamento ambiental = 30",
                    "Ocultamento de informações = 30",
                    "Omissão na manutenção de equipamentos = 40"
                ]
                motivo_combo.addItems(motivos)
                criterios_grid.addWidget(motivo_combo, 0, 1)

                # Consequência
                criterios_grid.addWidget(QLabel("2. Consequência ambiental:"), 1, 0)
                consequencia_combo = QComboBox()
                consequencias = [
                    "Selecione...",
                    "Potencial = 5",
                    "Desprezível = 10",
                    "Fraca = 20",
                    "Moderada = 30",
                    "Significativa = 50",
                    "Dano não reparável = 70"
                ]
                consequencia_combo.addItems(consequencias)
                criterios_grid.addWidget(consequencia_combo, 1, 1)

                # Saúde pública
                criterios_grid.addWidget(QLabel("3. Impacto à saúde pública:"), 2, 0)
                saude_combo = QComboBox()
                saude_opcoes = [
                    "Selecione...",
                    "Não houve = 0",
                    "Fraca = 5",
                    "Moderada = 10",
                    "Significativa = 20"
                ]
                saude_combo.addItems(saude_opcoes)
                criterios_grid.addWidget(saude_combo, 2, 1)

                quadro1_layout.addLayout(criterios_grid)

                # Pontuação
                pontuacao_frame = QFrame()
                pontuacao_frame.setStyleSheet("background-color: #e8f5e9; border-radius: 4px; padding: 8px;")
                pontuacao_layout = QHBoxLayout()
                pontuacao_layout.addWidget(QLabel("⭐ Pontuação Total:"))
                pontuacao_display = QLineEdit()
                pontuacao_display.setReadOnly(True)
                pontuacao_display.setStyleSheet("background-color: white; font-weight: bold; color: #006b3f;")
                pontuacao_display.setPlaceholderText("Aguardando seleção...")
                pontuacao_layout.addWidget(pontuacao_display)
                pontuacao_frame.setLayout(pontuacao_layout)
                quadro1_layout.addWidget(pontuacao_frame)

                quadro1_box.setLayout(quadro1_layout)
                main_layout.addWidget(quadro1_box)

                # -------- QUADRO 02 - VALORAÇÃO --------
                quadro2_box = QFrame()
                quadro2_box.setStyleSheet("""
                    QFrame {
                        background-color: #f8f9fa;
                        border-radius: 6px;
                        padding: 10px;
                    }
                """)
                quadro2_layout = QVBoxLayout()

                quadro2_title = QLabel("💰 QUADRO 02 - Valoração da Multa (ON 01/2024)")
                quadro2_title.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt; padding: 5px;")
                quadro2_layout.addWidget(quadro2_title)

                # Tipo de Infrator
                tipo_layout = QHBoxLayout()
                tipo_layout.addWidget(QLabel("Tipo de Infrator:"))
                tipo_infrator_combo = QComboBox()
                tipo_infrator_combo.addItems(["Pessoa Física", "Pessoa Jurídica"])
                tipo_layout.addWidget(tipo_infrator_combo)
                quadro2_layout.addLayout(tipo_layout)

                # Faixa de Receita
                faixa_layout = QHBoxLayout()
                faixa_layout.addWidget(QLabel("Faixa de Receita Bruta Mensal:"))
                faixa_receita_combo = QComboBox()
                faixas_pf = [
                    "Selecione a faixa...",
                    "Faixa A - Até 1 salário mínimo",
                    "Faixa B - 1 a 3 salários mínimos",
                    "Faixa C - 3 a 10 salários mínimos",
                    "Faixa D - 10 a 30 salários mínimos",
                    "Faixa E - 30 a 45 salários mínimos",
                    "Faixa F - Acima de 45 salários mínimos"
                ]
                faixas_pj = [
                    "Selecione a faixa...",
                    "Faixa A - MEI até R$ 81.000,00",
                    "Faixa B - ME até R$ 360.000,00",
                    "Faixa C - EPP até R$ 4.800.000,00",
                    "Faixa D - EMP até R$ 12.000.000,00",
                    "Faixa E - EGP-I até R$ 20.000.000,00",
                    "Faixa F - EGP-II acima de R$ 20.000.000,00"
                ]
                faixa_receita_combo.addItems(faixas_pf)

                def atualizar_faixas_receita_artigo():
                    faixa_receita_combo.clear()
                    if tipo_infrator_combo.currentText() == "Pessoa Jurídica":
                        faixa_receita_combo.addItems(faixas_pj)
                    else:
                        faixa_receita_combo.addItems(faixas_pf)

                tipo_infrator_combo.currentTextChanged.connect(atualizar_faixas_receita_artigo)

                faixa_layout.addWidget(faixa_receita_combo)
                quadro2_layout.addLayout(faixa_layout)

                percentual_label = QLabel("Percentual aplicado: -")
                percentual_label.setStyleSheet("color: #586069; font-style: italic; padding: 5px;")
                quadro2_layout.addWidget(percentual_label)

                quadro2_box.setLayout(quadro2_layout)
                main_layout.addWidget(quadro2_box)

                # -------- VALOR DA MULTA --------
                valor_frame = QFrame()
                valor_frame.setStyleSheet("background-color: #e3f2fd; border-radius: 6px; padding: 10px;")
                valor_layout = QHBoxLayout()
                valor_layout.addWidget(QLabel("💰 Valor da Multa:"))
                valor_input = QLineEdit()
                valor_input.setReadOnly(True)
                valor_input.setStyleSheet(
                    "background-color: white; font-weight: bold; color: #006b3f; font-size: 11pt;")
                valor_input.setPlaceholderText("Clique em 'Calcular Multa'")
                valor_layout.addWidget(valor_input)
                valor_frame.setLayout(valor_layout)
                main_layout.addWidget(valor_frame)

                # -------- BOTÕES --------
                botoes_layout = QHBoxLayout()
                botoes_layout.setSpacing(10)

                btn_calcular = QPushButton("🔢 Calcular Multa")
                btn_detalhes = QPushButton("ℹ️ Ver detalhes")
                btn_detalhes.setObjectName("secondary")

                def calcular_multa_artigo(checked=False,
                                          area=area_input,
                                          valor=valor_input,
                                          motivo=motivo_combo,
                                          consequencia=consequencia_combo,
                                          saude=saude_combo,
                                          pont=pontuacao_display,
                                          tipo_inf=tipo_infrator_combo,
                                          faixa=faixa_receita_combo,
                                          perc=percentual_label,
                                          chave_artigo=chave,
                                          dados_artigo=dados):

                    self.calcular_multa_com_dosimetria(
                        chave_artigo, area, valor, motivo, consequencia, saude,
                        pont, tipo_inf, faixa, perc,
                        dados_artigo.get("valor_base", 500),
                        dados_artigo.get("teto", 10000000)
                    )

                btn_calcular.clicked.connect(calcular_multa_artigo)
                btn_detalhes.clicked.connect(lambda: self.abrir_detalhes_artigo(chave))

                botoes_layout.addWidget(btn_calcular)
                botoes_layout.addWidget(btn_detalhes)
                botoes_layout.addStretch()

                main_layout.addLayout(botoes_layout)

                # Armazena os widgets do artigo
                self.artigos_inputs[chave] = {
                    "area": area_input,
                    "auto": auto_input,
                    "embargo": embargo_input,
                    "valor": valor_input,
                    "motivo": motivo_combo,
                    "consequencia": consequencia_combo,
                    "saude": saude_combo,
                    "pontuacao": pontuacao_display,
                    "tipo_infrator": tipo_infrator_combo,
                    "faixa_receita": faixa_receita_combo,
                    "percentual": percentual_label,
                    "titulo": dados.get("titulo", chave.upper()),
                    "eh_generico": False
                }

            else:
                # -------- ART. 48 - MULTAS SIMPLES --------
                valor_frame = QFrame()
                valor_frame.setStyleSheet("background-color: #e3f2fd; border-radius: 6px; padding: 10px;")
                valor_layout = QHBoxLayout()
                valor_layout.addWidget(QLabel("💰 Valor da Multa:"))
                valor_input = QLineEdit()
                valor_input.setReadOnly(True)
                valor_input.setStyleSheet(
                    "background-color: white; font-weight: bold; color: #006b3f; font-size: 11pt;")
                valor_input.setPlaceholderText("Clique em 'Calcular Multa'")
                valor_layout.addWidget(valor_input)
                valor_frame.setLayout(valor_layout)
                main_layout.addWidget(valor_frame)

                botoes_layout = QHBoxLayout()
                botoes_layout.setSpacing(10)

                btn_calcular = QPushButton("🔢 Calcular Multa")
                btn_detalhes = QPushButton("ℹ️ Ver detalhes")
                btn_detalhes.setObjectName("secondary")

                area_ref = area_input
                valor_ref = valor_input
                chave_ref = chave
                valor_ha_ref = dados["valor_por_ha"]

                btn_calcular.clicked.connect(
                    lambda: self.calcular_multa_artigo_simples(area_ref, valor_ref, valor_ha_ref))
                btn_detalhes.clicked.connect(lambda: self.abrir_detalhes_artigo(chave_ref))

                botoes_layout.addWidget(btn_calcular)
                botoes_layout.addWidget(btn_detalhes)
                botoes_layout.addStretch()

                main_layout.addLayout(botoes_layout)

                self.artigos_inputs[chave] = {
                    "area": area_input,
                    "auto": auto_input,
                    "embargo": embargo_input,
                    "valor": valor_input,
                    "titulo": dados.get("titulo", chave.upper()),
                    "eh_generico": False
                }

            grupo.setLayout(main_layout)
            layout.addWidget(grupo)

        # ==========================================================
        # ==========================================================
        # ARTIGO GENÉRICO - USUÁRIO DEFINE TUDO
        # ==========================================================
        # ==========================================================

        grupo_generico = QGroupBox("📌 ARTIGO GENÉRICO (Personalizado)")
        grupo_generico.setStyleSheet("""
            QGroupBox {
                background-color: #f0f7ff;
                border: 2px solid #006b3f;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 15px;
            }
            QGroupBox::title {
                font-weight: 700;
                color: #006b3f;
                font-size: 12pt;
            }
        """)

        generico_layout = QVBoxLayout()
        generico_layout.setSpacing(10)
        generico_layout.setContentsMargins(15, 10, 15, 15)

        # ==========================================================
        # DADOS DO ARTIGO GENÉRICO
        # ==========================================================

        grid_generico = QGridLayout()
        grid_generico.setVerticalSpacing(8)
        grid_generico.setHorizontalSpacing(15)

        # Nome do Artigo
        lbl_nome = QLabel("📌 Nome do Artigo:")
        lbl_nome.setStyleSheet("font-weight: 600; color: #006b3f;")
        nome_artigo_input = QLineEdit()
        nome_artigo_input.setPlaceholderText("Ex: Art. 80 - Descumprimento de embargo")
        nome_artigo_input.setMinimumWidth(300)
        grid_generico.addWidget(lbl_nome, 0, 0)
        grid_generico.addWidget(nome_artigo_input, 0, 1)

        # ==========================================================
        # VALORES MÍNIMO E MÁXIMO DEFINIDOS PELO USUÁRIO
        # ==========================================================

        # Valor Mínimo
        lbl_min = QLabel("💰 Valor Mínimo (R$):")
        lbl_min.setStyleSheet("font-weight: 500;")
        valor_min_input = QLineEdit()
        valor_min_input.setPlaceholderText("Ex: 5000,00")
        valor_min_input.setText("5000,00")
        grid_generico.addWidget(lbl_min, 1, 0)
        grid_generico.addWidget(valor_min_input, 1, 1)

        # Valor Máximo
        lbl_max = QLabel("💰 Valor Máximo (R$):")
        lbl_max.setStyleSheet("font-weight: 500;")
        valor_max_input = QLineEdit()
        valor_max_input.setPlaceholderText("Ex: 50000,00")
        valor_max_input.setText("50000,00")
        grid_generico.addWidget(lbl_max, 2, 0)
        grid_generico.addWidget(valor_max_input, 2, 1)

        generico_layout.addLayout(grid_generico)

        # ==========================================================
        # DESCRIÇÃO DA INFRAÇÃO (para o preview)
        # ==========================================================
        lbl_descricao_gen = QLabel("📝 Descrição da Infração:")
        lbl_descricao_gen.setStyleSheet("font-weight: 500;")
        descricao_gen_input = QTextEdit()
        descricao_gen_input.setPlaceholderText("Ex: descumprir embargo instituído por autoridade ambiental competente")
        descricao_gen_input.setMaximumHeight(60)

        grid_generico.addWidget(lbl_descricao_gen, 3, 0)
        grid_generico.addWidget(descricao_gen_input, 3, 1)

        # ==========================================================
        # CHECKBOX - ÁREA INFLUENCIA OU NÃO
        # ==========================================================

        check_area_influencia = QCheckBox("📐 A área influencia no cálculo da multa")
        check_area_influencia.setStyleSheet("""
            QCheckBox {
                font-weight: 500;
                padding: 5px;
            }
        """)
        check_area_influencia.setChecked(True)
        generico_layout.addWidget(check_area_influencia)

        # ==========================================================
        # CAMPO ÁREA - VISÍVEL APENAS SE CHECKED
        # ==========================================================

        area_generico_frame = QFrame()
        area_generico_layout = QHBoxLayout()
        area_generico_layout.setContentsMargins(20, 0, 0, 0)

        lbl_area_generico = QLabel("📐 Área (ha):")
        lbl_area_generico.setStyleSheet("font-weight: 500;")
        area_generico_input = QLineEdit()
        area_generico_input.setPlaceholderText("Digite a área em hectares")
        area_generico_input.setMinimumWidth(150)

        area_generico_layout.addWidget(lbl_area_generico)
        area_generico_layout.addWidget(area_generico_input)
        area_generico_layout.addStretch()
        area_generico_frame.setLayout(area_generico_layout)
        generico_layout.addWidget(area_generico_frame)

        # ==========================================================
        # FUNÇÃO PARA MOSTRAR/ESCONDER ÁREA
        # ==========================================================

        def toggle_area_generico(checked):
            area_generico_frame.setVisible(checked)

        check_area_influencia.toggled.connect(toggle_area_generico)

        # ==========================================================
        # AUTO E EMBARGO
        # ==========================================================

        grid_auto_embargo = QGridLayout()
        grid_auto_embargo.setVerticalSpacing(8)
        grid_auto_embargo.setHorizontalSpacing(15)

        lbl_auto_generico = QLabel("📄 Nº Auto de Infração:")
        lbl_auto_generico.setStyleSheet("font-weight: 500;")
        auto_generico_input = QLineEdit()
        auto_generico_input.setPlaceholderText("Número do auto")
        grid_auto_embargo.addWidget(lbl_auto_generico, 0, 0)
        grid_auto_embargo.addWidget(auto_generico_input, 0, 1)

        lbl_embargo_generico = QLabel("🚫 Nº Embargo:")
        lbl_embargo_generico.setStyleSheet("font-weight: 500;")
        embargo_generico_input = QLineEdit()
        embargo_generico_input.setPlaceholderText("Número do embargo")
        grid_auto_embargo.addWidget(lbl_embargo_generico, 1, 0)
        grid_auto_embargo.addWidget(embargo_generico_input, 1, 1)

        generico_layout.addLayout(grid_auto_embargo)

        # ==========================================================
        # SEPARADOR
        # ==========================================================

        separator_gen = QFrame()
        separator_gen.setFrameShape(QFrame.HLine)
        separator_gen.setStyleSheet("background-color: #e1e4e8; max-height: 1px;")
        generico_layout.addWidget(separator_gen)

        # ==========================================================
        # QUADRO 01 - DOSIMETRIA (GENÉRICO)
        # ==========================================================

        quadro1_gen_box = QFrame()
        quadro1_gen_box.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border-radius: 6px;
                padding: 10px;
                border: 1px solid #e1e4e8;
            }
        """)
        quadro1_gen_layout = QVBoxLayout()

        quadro1_gen_title = QLabel("📊 QUADRO 01 - Dosimetria da Infração (ON 01/2024)")
        quadro1_gen_title.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt; padding: 5px;")
        quadro1_gen_layout.addWidget(quadro1_gen_title)

        criterios_gen_grid = QGridLayout()
        criterios_gen_grid.setVerticalSpacing(10)

        # Motivo
        criterios_gen_grid.addWidget(QLabel("1. Motivo da Infração:"), 0, 0)
        motivo_gen_combo = QComboBox()
        motivos = [
            "Selecione...",
            "Não intencional = 5",
            "Intencional = 10",
            "Obtenção de vantagem pecuniária = 15",
            "Omissão deliberada em condicionantes = 20",
            "Burla ao licenciamento ambiental = 30",
            "Ocultamento de informações = 30",
            "Omissão na manutenção de equipamentos = 40"
        ]
        motivo_gen_combo.addItems(motivos)
        criterios_gen_grid.addWidget(motivo_gen_combo, 0, 1)

        # Consequência
        criterios_gen_grid.addWidget(QLabel("2. Consequência ambiental:"), 1, 0)
        consequencia_gen_combo = QComboBox()
        consequencias = [
            "Selecione...",
            "Potencial = 5",
            "Desprezível = 10",
            "Fraca = 20",
            "Moderada = 30",
            "Significativa = 50",
            "Dano não reparável = 70"
        ]
        consequencia_gen_combo.addItems(consequencias)
        criterios_gen_grid.addWidget(consequencia_gen_combo, 1, 1)

        # Saúde pública
        criterios_gen_grid.addWidget(QLabel("3. Impacto à saúde pública:"), 2, 0)
        saude_gen_combo = QComboBox()
        saude_opcoes = [
            "Selecione...",
            "Não houve = 0",
            "Fraca = 5",
            "Moderada = 10",
            "Significativa = 20"
        ]
        saude_gen_combo.addItems(saude_opcoes)
        criterios_gen_grid.addWidget(saude_gen_combo, 2, 1)

        quadro1_gen_layout.addLayout(criterios_gen_grid)

        # Pontuação
        pontuacao_gen_frame = QFrame()
        pontuacao_gen_frame.setStyleSheet("background-color: #e8f5e9; border-radius: 4px; padding: 8px;")
        pontuacao_gen_layout = QHBoxLayout()
        pontuacao_gen_layout.addWidget(QLabel("⭐ Pontuação Total:"))
        pontuacao_gen_display = QLineEdit()
        pontuacao_gen_display.setReadOnly(True)
        pontuacao_gen_display.setStyleSheet("background-color: white; font-weight: bold; color: #006b3f;")
        pontuacao_gen_display.setPlaceholderText("Aguardando seleção...")
        pontuacao_gen_layout.addWidget(pontuacao_gen_display)
        pontuacao_gen_frame.setLayout(pontuacao_gen_layout)
        quadro1_gen_layout.addWidget(pontuacao_gen_frame)

        quadro1_gen_box.setLayout(quadro1_gen_layout)
        generico_layout.addWidget(quadro1_gen_box)

        # ==========================================================
        # QUADRO 02 - VALORAÇÃO (GENÉRICO)
        # ==========================================================

        quadro2_gen_box = QFrame()
        quadro2_gen_box.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        quadro2_gen_layout = QVBoxLayout()

        quadro2_gen_title = QLabel("💰 QUADRO 02 - Valoração da Multa (ON 01/2024)")
        quadro2_gen_title.setStyleSheet("font-weight: bold; color: #006b3f; font-size: 11pt; padding: 5px;")
        quadro2_gen_layout.addWidget(quadro2_gen_title)

        # Tipo de Infrator
        tipo_gen_layout = QHBoxLayout()
        tipo_gen_layout.addWidget(QLabel("Tipo de Infrator:"))
        tipo_gen_combo = QComboBox()
        tipo_gen_combo.addItems(["Pessoa Física", "Pessoa Jurídica"])
        tipo_gen_layout.addWidget(tipo_gen_combo)
        quadro2_gen_layout.addLayout(tipo_gen_layout)

        # Faixa de Receita
        faixa_gen_layout = QHBoxLayout()
        faixa_gen_layout.addWidget(QLabel("Faixa de Receita Bruta Mensal:"))
        faixa_gen_combo = QComboBox()
        faixas_pf = [
            "Selecione a faixa...",
            "Faixa A - Até 1 salário mínimo",
            "Faixa B - 1 a 3 salários mínimos",
            "Faixa C - 3 a 10 salários mínimos",
            "Faixa D - 10 a 30 salários mínimos",
            "Faixa E - 30 a 45 salários mínimos",
            "Faixa F - Acima de 45 salários mínimos"
        ]
        faixas_pj = [
            "Selecione a faixa...",
            "Faixa A - MEI até R$ 81.000,00",
            "Faixa B - ME até R$ 360.000,00",
            "Faixa C - EPP até R$ 4.800.000,00",
            "Faixa D - EMP até R$ 12.000.000,00",
            "Faixa E - EGP-I até R$ 20.000.000,00",
            "Faixa F - EGP-II acima de R$ 20.000.000,00"
        ]
        faixa_gen_combo.addItems(faixas_pf)

        def atualizar_faixas_receita_gen():
            faixa_gen_combo.clear()
            if tipo_gen_combo.currentText() == "Pessoa Jurídica":
                faixa_gen_combo.addItems(faixas_pj)
            else:
                faixa_gen_combo.addItems(faixas_pf)

        tipo_gen_combo.currentTextChanged.connect(atualizar_faixas_receita_gen)

        faixa_gen_layout.addWidget(faixa_gen_combo)
        quadro2_gen_layout.addLayout(faixa_gen_layout)

        percentual_gen_label = QLabel("Percentual aplicado: -")
        percentual_gen_label.setStyleSheet("color: #586069; font-style: italic; padding: 5px;")
        quadro2_gen_layout.addWidget(percentual_gen_label)

        quadro2_gen_box.setLayout(quadro2_gen_layout)
        generico_layout.addWidget(quadro2_gen_box)

        # ==========================================================
        # VALOR DA MULTA (GENÉRICO)
        # ==========================================================

        valor_gen_frame = QFrame()
        valor_gen_frame.setStyleSheet("background-color: #e3f2fd; border-radius: 6px; padding: 10px;")
        valor_gen_layout = QHBoxLayout()
        valor_gen_layout.addWidget(QLabel("💰 Valor da Multa:"))
        valor_gen_input = QLineEdit()
        valor_gen_input.setReadOnly(True)
        valor_gen_input.setStyleSheet(
            "background-color: white; font-weight: bold; color: #006b3f; font-size: 11pt;")
        valor_gen_input.setPlaceholderText("Clique em 'Calcular Multa'")
        valor_gen_layout.addWidget(valor_gen_input)
        valor_gen_frame.setLayout(valor_gen_layout)
        generico_layout.addWidget(valor_gen_frame)

        # ==========================================================
        # BOTÕES (GENÉRICO)
        # ==========================================================

        botoes_gen_layout = QHBoxLayout()
        botoes_gen_layout.setSpacing(10)

        btn_calcular_gen = QPushButton("🔢 Calcular Multa")
        btn_detalhes_gen = QPushButton("ℹ️ Ver detalhes")
        btn_detalhes_gen.setObjectName("secondary")

        # ==========================================================
        # FUNÇÃO CALCULAR MULTA GENÉRICA
        # ==========================================================

        def calcular_multa_generico():
            try:
                # Obtém o nome do artigo
                nome_artigo = nome_artigo_input.text().strip()
                if not nome_artigo:
                    QMessageBox.warning(self, "Aviso", "Digite o nome do artigo!")
                    return

                # Obtém valores mínimo e máximo
                try:
                    valor_min = float(valor_min_input.text().replace(',', '.').strip())
                    valor_max = float(valor_max_input.text().replace(',', '.').strip())
                except ValueError:
                    QMessageBox.warning(self, "Aviso", "Valores mínimo e máximo devem ser números válidos!")
                    return

                if valor_min >= valor_max:
                    QMessageBox.warning(self, "Aviso", "O valor mínimo deve ser menor que o valor máximo!")
                    return

                # Verifica se a área influencia
                area_influencia = check_area_influencia.isChecked()

                # Obtém a área
                area_original = 1.0
                try:
                    if area_influencia:
                        area_texto = area_generico_input.text().strip()
                        if area_texto:
                            area_original = float(area_texto.replace(',', '.'))
                        if area_original <= 0:
                            QMessageBox.warning(self, "Aviso", "A área deve ser maior que zero!")
                            return
                except ValueError:
                    QMessageBox.warning(self, "Aviso", "Digite um valor válido para a área!")
                    return

                # Arredonda a área para cima se influencia
                if area_influencia:
                    area = math.ceil(area_original)
                    if area != area_original:
                        resposta = QMessageBox.question(self, "Arredondamento",
                                                        f"Área original: {area_original} ha\nÁrea arredondada para cima: {area} ha\n\nDeseja continuar?",
                                                        QMessageBox.Yes | QMessageBox.No)
                        if resposta == QMessageBox.No:
                            return
                    else:
                        area = area_original
                else:
                    area = 1.0

                # ==========================================================
                # CALCULA A PONTUAÇÃO
                # ==========================================================
                motivo_pontos = 0
                consequencia_pontos = 0
                saude_pontos = 0

                motivo_text = motivo_gen_combo.currentText()
                if "=" in motivo_text and motivo_text != "Selecione...":
                    try:
                        motivo_pontos = int(motivo_text.split("=")[1].strip())
                    except:
                        motivo_pontos = 0

                consequencia_text = consequencia_gen_combo.currentText()
                if "=" in consequencia_text and consequencia_text != "Selecione...":
                    try:
                        consequencia_pontos = int(consequencia_text.split("=")[1].strip())
                    except:
                        consequencia_pontos = 0

                saude_text = saude_gen_combo.currentText()
                if "=" in saude_text and saude_text != "Selecione...":
                    try:
                        saude_pontos = int(saude_text.split("=")[1].strip())
                    except:
                        saude_pontos = 0

                pontuacao_total = motivo_pontos + consequencia_pontos + saude_pontos

                # Define o nível
                if 10 <= pontuacao_total <= 20:
                    nivel = "Nível A (Gravidade Baixa)"
                elif 21 <= pontuacao_total <= 40:
                    nivel = "Nível B (Gravidade Média-Baixa)"
                elif 41 <= pontuacao_total <= 60:
                    nivel = "Nível C (Gravidade Média)"
                elif 61 <= pontuacao_total <= 80:
                    nivel = "Nível D (Gravidade Média-Alta)"
                elif 81 <= pontuacao_total <= 100:
                    nivel = "Nível E (Gravidade Alta)"
                else:
                    nivel = ""

                if pontuacao_total >= 10:
                    pontuacao_gen_display.setText(f"{pontuacao_total} pontos - {nivel}")
                    pontuacao_gen_display.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold;")
                else:
                    QMessageBox.warning(self, "Aviso", f"Pontuação mínima é 10 pontos. Atual: {pontuacao_total}")
                    return

                # ==========================================================
                # OBTÉM FAIXA E PERCENTUAL
                # ==========================================================
                faixa_texto = faixa_gen_combo.currentText()
                if "Selecione" in faixa_texto or not faixa_texto:
                    QMessageBox.warning(self, "Aviso", "Selecione a faixa de receita bruta mensal.")
                    return

                faixa = faixa_texto.split(" - ")[0]
                tipo_infrator = tipo_gen_combo.currentText()
                nivel_curto = nivel.split(" (")[0] if " (" in nivel else nivel

                if tipo_infrator == "Pessoa Jurídica":
                    perc_min, perc_max = self.get_percentual_pj(faixa, nivel_curto)
                else:
                    perc_min, perc_max = self.get_percentual_por_faixa_e_nivel(faixa, nivel_curto)

                # ==========================================================
                # CALCULA A MULTA USANDO O INTERVALO DO USUÁRIO
                # ==========================================================
                # Aplica o percentual sobre o teto (valor máximo) e soma ao mínimo
                acrescimo_min = valor_max * (perc_min / 100)
                acrescimo_max = valor_max * (perc_max / 100)

                valor_multa_min = valor_min + acrescimo_min
                valor_multa_max = valor_min + acrescimo_max

                # Se área influencia, multiplica
                if area_influencia:
                    valor_multa_min *= area
                    valor_multa_max *= area

                # Aplica dobro se UC marcado
                if self.is_uc_marcado():
                    valor_multa_min *= 2
                    valor_multa_max *= 2

                # ==========================================================
                # FORMATA E EXIBE
                # ==========================================================
                min_str = self.gerenciador._formatar_moeda_br(valor_multa_min)
                max_str = self.gerenciador._formatar_moeda_br(valor_multa_max)

                uc_texto = " (DOBRO - Art. 93 - UC)" if self.is_uc_marcado() else ""
                area_texto_extra = f" | Área: {area} ha" if area_influencia else ""

                percentual_gen_label.setText(
                    f"✓ {nome_artigo} | Nível: {nivel_curto} | Faixa: {faixa} | Percentual: {perc_min:.3f}% a {perc_max:.3f}%{uc_texto}{area_texto_extra}"
                )

                valor_gen_input.setText(f"{min_str} a {max_str}")
                valor_gen_input.setStyleSheet(
                    "background-color: #d4edda; color: #006b3f; font-weight: bold; font-size: 11pt;")

                # ==========================================================
                # SALVA OS DADOS NO DICIONÁRIO
                # ==========================================================
                self.artigos_inputs["art_generico"] = {
                    "area": area_generico_input,
                    "auto": auto_generico_input,
                    "embargo": embargo_generico_input,
                    "valor": valor_gen_input,
                    "motivo": motivo_gen_combo,
                    "consequencia": consequencia_gen_combo,
                    "saude": saude_gen_combo,
                    "pontuacao": pontuacao_gen_display,
                    "tipo_infrator": tipo_gen_combo,
                    "faixa_receita": faixa_gen_combo,
                    "percentual": percentual_gen_label,
                    "titulo": nome_artigo,
                    "eh_generico": True,
                    "valor_min": valor_min,
                    "valor_max": valor_max,
                    "area_influencia": area_influencia,
                    "descricao_infracao": descricao_gen_input,
                    "check_area": check_area_influencia
                }

                # Atualiza a pré-visualização
                self.atualizar_preview()

            except Exception as e:
                QMessageBox.critical(self, "Erro inesperado", f"Erro ao calcular: {str(e)}")

        btn_calcular_gen.clicked.connect(calcular_multa_generico)
        btn_detalhes_gen.clicked.connect(lambda: self.abrir_detalhes_artigo("art_generico"))

        botoes_gen_layout.addWidget(btn_calcular_gen)
        botoes_gen_layout.addWidget(btn_detalhes_gen)
        botoes_gen_layout.addStretch()

        generico_layout.addLayout(botoes_gen_layout)

        grupo_generico.setLayout(generico_layout)
        layout.addWidget(grupo_generico)

        # ==========================================================
        # CHECKBOXES DOS ARTIGOS APLICÁVEIS
        # ==========================================================

        grupo_check = QGroupBox("📍 Artigos Aplicáveis")
        grupo_check.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 1px solid #e1e4e8;
                border-radius: 8px;
            }
        """)

        check_layout = QHBoxLayout()
        check_layout.setSpacing(20)
        check_layout.setContentsMargins(15, 10, 15, 10)

        self.check_art48 = QCheckBox("Art. 48")
        self.check_art66 = QCheckBox("Art. 66")
        self.check_art79 = QCheckBox("Art. 79")
        self.check_art_generico = QCheckBox("Art. Genérico")

        self.check_art48.stateChanged.connect(self.atualizar_preview)
        self.check_art66.stateChanged.connect(self.atualizar_preview)
        self.check_art79.stateChanged.connect(self.atualizar_preview)
        self.check_art_generico.stateChanged.connect(self.atualizar_preview)

        for cb in [self.check_art48, self.check_art66, self.check_art79, self.check_art_generico]:
            if cb is not None:
                cb.setStyleSheet("""
                    QCheckBox {
                        spacing: 8px;
                        font-weight: 500;
                    }
                """)

        check_layout.addWidget(self.check_art48)
        check_layout.addWidget(self.check_art66)
        check_layout.addWidget(self.check_art79)
        check_layout.addWidget(self.check_art_generico)
        check_layout.addStretch()

        grupo_check.setLayout(check_layout)
        layout.addWidget(grupo_check)
        layout.addStretch()

        container.setLayout(layout)
        scroll.setWidget(container)

        final_layout = QVBoxLayout()
        final_layout.setContentsMargins(0, 0, 0, 0)
        final_layout.addWidget(scroll)
        widget.setLayout(final_layout)

        return widget
    
    def calcular_multa_artigo_simples(self, area_input, valor_input, valor_por_ha):
        try:
            if area_input is None:
                QMessageBox.warning(self, "Erro", "Campo de área não encontrado!")
                return
            
            texto_area = area_input.text().strip()
            if not texto_area:
                QMessageBox.warning(self, "Erro", "Por favor, digite um valor para a área!")
                return
            
            texto_area = texto_area.replace(",", ".")
            
            try:
                area_original = float(texto_area)
            except ValueError:
                QMessageBox.warning(self, "Erro de conversão", f"Valor inválido: '{texto_area}'")
                return
            
            if area_original <= 0:
                QMessageBox.warning(self, "Erro", "A área deve ser maior que zero!")
                return
            
            area = math.ceil(area_original)
            
            if area != area_original:
                resposta = QMessageBox.question(self, "Arredondamento de área",
                    f"Área original: {area_original} ha\nÁrea arredondada para cima: {area} ha\n\nDeseja continuar?",
                    QMessageBox.Yes | QMessageBox.No)
                if resposta == QMessageBox.No:
                    return
                area_input.setText(str(area))
            
            valor_total = area * valor_por_ha

            # VERIFICA SE UC ESTÁ MARCADO - APLICA DOBRO (Art. 93)
            if self.is_uc_marcado():
                valor_total *= 2

            valor_total_str = f"R$ {valor_total:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
            if self.is_uc_marcado():
                valor_total_str += " (DOBRO - Art. 93 - UC)"
                valor_input.setStyleSheet(
                    "background-color: #d4edda; color: #006b3f; font-weight: bold; font-size: 11pt; border: 2px solid #ff6600;")
            else:
                valor_input.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold; font-size: 11pt;")
            valor_input.setText(valor_total_str)

            self.atualizar_preview()

        except Exception as e:
            QMessageBox.critical(self, "Erro inesperado", f"Erro ao calcular: {str(e)}")

    def calcular_multa_com_dosimetria(self, chave, area_input, valor_input,
                                      motivo_combo, consequencia_combo, saude_combo,
                                      pontuacao_display, tipo_infrator_combo, faixa_receita_combo,
                                      percentual_label, valor_base, teto):
        """
        Calcula multa para artigos com dosimetria (Art. 66 e Art. 79)
        NÃO depende da área - é uma multa por ATIVIDADE
        """
        try:
            # ==========================================================
            # VERIFICA SE A ÁREA É OBRIGATÓRIA (Art. 66 e 79 NÃO PRECISAM)
            # ==========================================================
            # Para Art. 66 e Art. 79, a área é informativa, não influencia o cálculo
            area_texto = area_input.text().strip()
            area_original = 1.0  # Valor padrão (apenas para exibição)

            try:
                if area_texto:
                    area_original = float(area_texto.replace(',', '.'))
            except:
                area_original = 1.0

            # ==========================================================
            # CALCULA A PONTUAÇÃO
            # ==========================================================
            motivo_pontos = 0
            consequencia_pontos = 0
            saude_pontos = 0

            motivo_text = motivo_combo.currentText()
            if "=" in motivo_text and motivo_text != "Selecione...":
                try:
                    motivo_pontos = int(motivo_text.split("=")[1].strip())
                except:
                    motivo_pontos = 0

            consequencia_text = consequencia_combo.currentText()
            if "=" in consequencia_text and consequencia_text != "Selecione...":
                try:
                    consequencia_pontos = int(consequencia_text.split("=")[1].strip())
                except:
                    consequencia_pontos = 0

            saude_text = saude_combo.currentText()
            if "=" in saude_text and saude_text != "Selecione...":
                try:
                    saude_pontos = int(saude_text.split("=")[1].strip())
                except:
                    saude_pontos = 0

            pontuacao_total = motivo_pontos + consequencia_pontos + saude_pontos

            # Define o nível
            if 10 <= pontuacao_total <= 20:
                nivel = "Nível A (Gravidade Baixa)"
            elif 21 <= pontuacao_total <= 40:
                nivel = "Nível B (Gravidade Média-Baixa)"
            elif 41 <= pontuacao_total <= 60:
                nivel = "Nível C (Gravidade Média)"
            elif 61 <= pontuacao_total <= 80:
                nivel = "Nível D (Gravidade Média-Alta)"
            elif 81 <= pontuacao_total <= 100:
                nivel = "Nível E (Gravidade Alta)"
            else:
                nivel = ""

            if pontuacao_total >= 10:
                pontuacao_display.setText(f"{pontuacao_total} pontos - {nivel}")
                pontuacao_display.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold;")
            else:
                QMessageBox.warning(self, "Aviso", f"Pontuação mínima é 10 pontos. Atual: {pontuacao_total}")
                return

            # ==========================================================
            # DEFINE OS VALORES BASE E TETO PARA CADA ARTIGO
            # ==========================================================
            # Art. 66: R$ 500,00 a R$ 10.000.000,00
            # Art. 79: R$ 10.000,00 a R$ 10.000.000,00
            if chave == "art66":
                valor_base_artigo = 500.00
                teto_artigo = 10000000.00
                nome_artigo = "Art. 66"
            elif chave == "art79":
                valor_base_artigo = 10000.00
                teto_artigo = 10000000.00
                nome_artigo = "Art. 79"
            else:
                valor_base_artigo = valor_base
                teto_artigo = teto

            # ==========================================================
            # OBTÉM FAIXA E PERCENTUAL
            # ==========================================================
            faixa_texto = faixa_receita_combo.currentText()
            if "Selecione" in faixa_texto or not faixa_texto:
                QMessageBox.warning(self, "Aviso", "Selecione a faixa de receita bruta mensal.")
                return

            faixa = faixa_texto.split(" - ")[0]
            tipo_infrator = tipo_infrator_combo.currentText()

            # Extrai o Nível curto
            nivel_curto = nivel.split(" (")[0] if " (" in nivel else nivel

            # Obtém percentuais
            if tipo_infrator == "Pessoa Jurídica":
                perc_min, perc_max = self.get_percentual_pj(faixa, nivel_curto)
            else:
                perc_min, perc_max = self.get_percentual_por_faixa_e_nivel(faixa, nivel_curto)

            # ==========================================================
            # CALCULA A MULTA (NÃO DEPENDE DA ÁREA)
            # ==========================================================
            # Fórmula: valor_base + (percentual * teto)
            acrescimo_min = teto_artigo * (perc_min / 100)
            acrescimo_max = teto_artigo * (perc_max / 100)

            valor_multa_min = valor_base_artigo + acrescimo_min
            valor_multa_max = valor_base_artigo + acrescimo_max

            # ==========================================================
            # APLICA DOBRO SE UC ESTIVER MARCADO
            # ==========================================================
            if self.is_uc_marcado():
                valor_multa_min *= 2
                valor_multa_max *= 2

            # ==========================================================
            # FORMATA OS VALORES
            # ==========================================================
            valor_min_str = self.gerenciador._formatar_moeda_br(valor_multa_min)
            valor_max_str = self.gerenciador._formatar_moeda_br(valor_multa_max)

            # ==========================================================
            # EXIBE O RESULTADO
            # ==========================================================
            uc_texto = " (DOBRO - Art. 93 - UC)" if self.is_uc_marcado() else ""

            percentual_label.setText(
                f"✓ {nome_artigo} | Nível: {nivel_curto} | Faixa: {faixa} | Percentual: {perc_min:.3f}% a {perc_max:.3f}%{uc_texto}"
            )

            # Mostra o valor calculado (multa por atividade, não por área)
            valor_input.setText(f"{valor_min_str} a {valor_max_str}")
            valor_input.setStyleSheet("background-color: #d4edda; color: #006b3f; font-weight: bold; font-size: 11pt;")

            # ==========================================================
            # ATUALIZA A PRÉ-VISUALIZAÇÃO
            # ==========================================================
            self.atualizar_preview()

        except Exception as e:
            QMessageBox.critical(self, "Erro inesperado", f"Erro ao calcular: {str(e)}")

    def abrir_detalhes_artigo(self, tipo):
        conteudos = {
            "art48": "Art. 48 do Decreto 6514/2008\n\nImpedir ou dificultar a regeneração natural de florestas e demais formas de vegetação.\n\nMulta: R$ 5.000,00 por hectare.",
            "art66": "Art. 66 do Decreto 6514/2008\n\nSuprimir vegetação nativa em Reserva Legal.\n\nMulta calculada por dosimetria conforme ON 01/2024.",
            "art79": "Art. 79 do Decreto 6514/2008\n\nDescumprir embargo instituído por autoridade ambiental.\n\nMulta calculada por dosimetria conforme ON 01/2024."
        }
        titulos = {"art48": "Art. 48", "art66": "Art. 66", "art79": "Art. 79"}
        modal = ModalDetalhes(titulos[tipo], conteudos[tipo], self)
        modal.exec()

    def criar_aba_novos_modelos(self):
        """Cria a aba com os relatórios específicos"""
        widget = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)

        # Seção de seleção dos relatórios
        grupo_selecao = QGroupBox("📄 SELECIONE OS RELATÓRIOS PARA GERAR")
        grupo_selecao.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 2px solid #006b3f;
                border-radius: 8px;
                margin-top: 15px;
            }
            QGroupBox::title {
                color: #006b3f;
                font-weight: bold;
                font-size: 12pt;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
            }
        """)

        selecao_layout = QVBoxLayout()
        selecao_layout.setSpacing(10)

        self.check_relatorio_ina = QCheckBox("📊 Relatório Inã - Supressão de vegetação nativa")
        self.check_despacho = QCheckBox("📝 Despacho de Autocomposição")
        self.check_minuta = QCheckBox("📄 Modelo de Minuta de Auto de Infração")
        self.check_autodenuncia = QCheckBox("📋 Relatório INÃ - Autodenúncia")

        # NOVOS CHECKBOXES
        self.check_barramento = QCheckBox("💧 Relatório INÃ - Barramento")
        self.check_parcelamento = QCheckBox("🏘️ Relatório INÃ - Parcelamento")

        for cb in [self.check_relatorio_ina, self.check_despacho, self.check_minuta,
                   self.check_autodenuncia, self.check_barramento, self.check_parcelamento]:
            cb.setStyleSheet("""
                QCheckBox {
                    font-size: 11pt;
                    padding: 8px;
                    font-weight: 500;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                }
            """)
            selecao_layout.addWidget(cb)

        grupo_selecao.setLayout(selecao_layout)
        layout.addWidget(grupo_selecao)

        # ... resto do método permanece igual ...
        
        # Botões de ação
        grupo_acoes = QGroupBox("⚙️ AÇÕES")
        grupo_acoes.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 1px solid #e1e4e8;
                border-radius: 8px;
                margin-top: 10px;
            }
            QGroupBox::title {
                color: #006b3f;
                font-weight: bold;
            }
        """)
        
        acoes_layout = QVBoxLayout()
        acoes_layout.setSpacing(10)
        
        self.btn_gerar_novos = QPushButton("✅ GERAR RELATÓRIOS SELECIONADOS")
        self.btn_gerar_novos.setStyleSheet("""
            QPushButton {
                background-color: #006b3f;
                font-size: 12pt;
                padding: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #008c52;
            }
        """)
        self.btn_gerar_novos.clicked.connect(self.gerar_novos_relatorios)
        
        self.progress_bar_novos = QProgressBar()
        self.progress_bar_novos.setVisible(False)
        
        self.label_status_novos = QLabel("Aguardando ação...")
        self.label_status_novos.setStyleSheet("color: #586069; padding: 5px;")
        
        acoes_layout.addWidget(self.btn_gerar_novos)
        acoes_layout.addWidget(self.progress_bar_novos)
        acoes_layout.addWidget(self.label_status_novos)
        
        grupo_acoes.setLayout(acoes_layout)
        layout.addWidget(grupo_acoes)
        
        # Informação adicional
        info_label = QLabel("💡 Dica: Os dados para geração dos relatórios são coletados das abas 'Dados Gerais', 'Autos e Embargos' e 'Artigos'")
        info_label.setStyleSheet("color: #586069; font-style: italic; padding: 10px; background-color: #f8f9fa; border-radius: 6px;")
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        layout.addStretch()
        
        container.setLayout(layout)
        scroll.setWidget(container)
        
        final_layout = QVBoxLayout()
        final_layout.addWidget(scroll)
        widget.setLayout(final_layout)
        
        return widget

    def criar_aba_creditos(self):
        """Cria a aba de créditos dos desenvolvedores"""
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Card principal
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 12px;
                padding: 20px;
            }
        """)
        card_layout = QVBoxLayout()
        
        # Título
        titulo = QLabel("👥 EQUIPE DE DESENVOLVIMENTO")
        titulo.setStyleSheet("""
            font-size: 18pt;
            font-weight: bold;
            color: #006b3f;
            padding: 10px;
            border-bottom: 2px solid #006b3f;
        """)
        titulo.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(titulo)
        
        # Subtítulo
        subtitulo = QLabel("Gerência de Geoprocessamento e Sensoriamento Remoto - GEGEO")
        subtitulo.setStyleSheet("""
            font-size: 12pt;
            color: #586069;
            padding: 5px;
            margin-bottom: 20px;
        """)
        subtitulo.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(subtitulo)
        
        # Lista de desenvolvedores
        desenvolvedores = [
            ("👩‍💼 Raissa Daher Alves", "Analista Ambiental", "raissa.alves@go.gov.br"),
            ("👨‍💼 Joelcio Cláudio Lima", "Analista Ambiental", "joelcio.lima@go.gov.br"),
            ("👨‍💻 Vicente de Paula Sousa Júnior", "Analista Ambiental", "vicente.junior@go.gov.br"),
        ]
        
        for nome, cargo, email in desenvolvedores:
            dev_frame = QFrame()
            dev_frame.setStyleSheet("""
                QFrame {
                    background-color: #f8f9fa;
                    border-radius: 8px;
                    padding: 10px;
                    margin: 5px;
                }
            """)
            dev_layout = QHBoxLayout()
            
            # Ícone e nome
            nome_label = QLabel(f"{nome}")
            nome_label.setStyleSheet("font-size: 12pt; font-weight: bold; color: #2c3e50;")
            nome_label.setMinimumWidth(250)
            
            cargo_label = QLabel(cargo)
            cargo_label.setStyleSheet("color: #586069;")
            cargo_label.setMinimumWidth(150)
            
            email_label = QLabel(email)
            email_label.setStyleSheet("color: #006b3f; font-style: italic;")
            
            dev_layout.addWidget(nome_label)
            dev_layout.addWidget(cargo_label)
            dev_layout.addWidget(email_label)
            dev_layout.addStretch()
            
            dev_frame.setLayout(dev_layout)
            card_layout.addWidget(dev_frame)
        
        # Informações do sistema
        info_frame = QFrame()
        info_frame.setStyleSheet("""
            QFrame {
                background-color: #e8f5e9;
                border-radius: 8px;
                padding: 15px;
                margin-top: 20px;
            }
        """)
        info_layout = QVBoxLayout()
        
        versao_label = QLabel("📌 Versão: 2.0.0")
        versao_label.setStyleSheet("font-weight: bold; color: #006b3f;")
        
        data_label = QLabel(f"📅 Última atualização: {datetime.now().strftime('%d/%m/%Y')}")
        
        info_layout.addWidget(versao_label)
        info_layout.addWidget(data_label)
        info_frame.setLayout(info_layout)
        card_layout.addWidget(info_frame)
        
        card.setLayout(card_layout)
        layout.addWidget(card)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget

    def atualizar_preview(self):
        """Atualiza a pré-visualização com fonte maior e títulos claros"""
        # APENAS DESCRIÇÃO DAS INFRAÇÕES (já inclui áreas e artigos)
        descricao = self.obter_descricao_infracao()

        # VERIFICA SE DESCRICAO NÃO É NONE OU VAZIA
        if descricao is None:
            descricao = "Nenhuma infração selecionada para visualização."

        uc_status = "✅ ATIVADO - Multas em DOBRO (Art. 93)" if self.is_uc_marcado() else "❌ Desativado"
        uc_color = "#28a745" if self.is_uc_marcado() else "#dc3545"

        # Processa o texto para HTML
        descricao_html = descricao.replace('\n', '<br>')

        # Adiciona cor de destaque para valores de multa
        import re
        descricao_html = re.sub(r'(R\$[\s]*[\d\.,]+(?: a R\$[\s]*[\d\.,]+)?)',
                                r'<span style="color: #006b3f; font-weight: bold;">\1</span>', descricao_html)

        # Adiciona destaque para os títulos dos artigos
        descricao_html = re.sub(r'(┌─────────────────────────────────────────────────────────────────────────────┐)',
                                r'<span style="background-color: #e8f5e9; display: block; border-radius: 4px;">\1</span>',
                                descricao_html)

        # Destaca os títulos dentro dos cards
        descricao_html = re.sub(
            r'(📌 Art\. \d+ - [^\n]+)',
            r'<span style="font-size: 14pt; font-weight: bold; color: #006b3f;">\1</span>',
            descricao_html
        )

        # Destaca AUTO DE INFRAÇÃO e EMBARGO
        descricao_html = re.sub(
            r'(📄 AUTO DE INFRAÇÃO:)',
            r'<span style="font-size: 13pt; font-weight: bold; color: #006b3f;">\1</span>',
            descricao_html
        )
        descricao_html = re.sub(
            r'(🚫 EMBARGO:)',
            r'<span style="font-size: 13pt; font-weight: bold; color: #856404;">\1</span>',
            descricao_html
        )
        descricao_html = re.sub(
            r'(📍 DESCRIÇÃO DE ACESSO AO IMÓVEL:)',
            r'<span style="font-size: 13pt; font-weight: bold; color: #006b3f;">\1</span>',
            descricao_html
        )

        # Constrói o HTML completo
        preview = f"""<!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{
            font-family: 'Segoe UI', 'Roboto', Arial, sans-serif;
            font-size: 13pt;
            margin: 15px;
            background-color: #f8f9fa;
            color: #2c3e50;
            line-height: 1.8;
        }}
        .card {{
            background-color: white;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-left: 4px solid #006b3f;
        }}
        .titulo-principal {{
            font-size: 18pt;
            font-weight: bold;
            color: #006b3f;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #006b3f;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .uc-status {{
            border: 1px solid {uc_color};
            border-radius: 20px;
            padding: 8px 16px;
            display: inline-block;
            font-weight: 500;
            font-size: 13pt;
            color: {uc_color};
            background-color: transparent;
        }}
        pre {{
            white-space: pre-wrap;
            word-wrap: break-word;
            font-family: 'Courier New', 'Consolas', monospace;
            font-size: 13pt;
            line-height: 1.8;
            background-color: #fafafa;
            padding: 15px;
            border-radius: 6px;
            border: 1px solid #e1e4e8;
        }}
        .area-titulo {{
            font-size: 14pt;
            font-weight: bold;
            color: #006b3f;
            margin-top: 10px;
            margin-bottom: 5px;
        }}
        .sub-titulo {{
            font-size: 13pt;
            font-weight: bold;
            color: #006b3f;
        }}
    </style>
    </head>
    <body>

    <!-- CARD - STATUS UC -->
    <div class="card">
        <div class="titulo-principal">
            UNIDADE DE CONSERVAÇÃO (UC)
        </div>
        <div class="uc-status" style="display: inline-block;">
            {uc_status}
        </div>
        <div style="margin-top: 10px; color: #6c757d; font-size: 12pt;">
            {self._get_uc_explicacao()}
        </div>
    </div>

    <!-- CARD - TEXTO PARA INÃ FISCALIZAÇÃO -->
    <div class="card">
        <div class="titulo-principal">
            📝 TEXTO PARA INÃ FISCALIZAÇÃO
        </div>
        <pre>{descricao_html}</pre>
    </div>

    </body>
    </html>"""

        self.preview_text.setHtml(preview)

    def gerar_txt_preview(self):
        """Gera um arquivo .txt com o conteúdo da pré-visualização"""
        try:
            # Obtém o texto do preview
            texto_preview = self.preview_text.toPlainText()

            if not texto_preview or texto_preview.strip() == "":
                QMessageBox.warning(self, "Aviso", "Não há conteúdo para salvar. Atualize a pré-visualização primeiro.")
                return

            # Pergunta onde salvar
            pasta = QFileDialog.getExistingDirectory(self, "Escolha a pasta para salvar o arquivo TXT")
            if not pasta:
                return

            # Gera nome do arquivo
            imovel = self.inputs.get("imovel", QLineEdit()).text().strip() or "IMOVEL_NAO_INFORMADO"
            data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"Preview_INA_{imovel}_{data_atual}.txt"

            # Remove caracteres inválidos para nome de arquivo
            nome_arquivo = "".join(c for c in nome_arquivo if c.isalnum() or c in '._- ')

            caminho = os.path.join(pasta, nome_arquivo)

            # Salva o arquivo
            with open(caminho, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write("PREVIEW - TEXTO PARA INÃ FISCALIZAÇÃO\n")
                f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")
                f.write(texto_preview)
                f.write("\n\n" + "=" * 80 + "\n")
                f.write("FIM DO PREVIEW\n")

            QMessageBox.information(self, "Sucesso", f"Arquivo TXT gerado com sucesso!\n{caminho}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar arquivo TXT: {str(e)}")

    def gerar_relatorio_ina(self, dados):
        """Gera o conteúdo do Relatório Inã - Versão completa"""

        # Função para converter número em extenso
        def numero_extenso(valor):
            """Converte valor numérico para extenso (ex: 40100.00 -> quarenta mil e cem reais)"""
            try:
                valor_int = int(valor)
                centavos = int(round((valor - valor_int) * 100))

                unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
                especiais = {10: 'dez', 11: 'onze', 12: 'doze', 13: 'treze', 14: 'quatorze',
                             15: 'quinze', 16: 'dezesseis', 17: 'dezessete', 18: 'dezoito', 19: 'dezenove'}
                dezenas = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta',
                           'sessenta', 'setenta', 'oitenta', 'noventa']
                centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos',
                            'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']

                def converter_ate_999(n):
                    if n == 0:
                        return ''
                    if n == 100:
                        return 'cem'

                    texto = ''
                    if n >= 100:
                        texto += centenas[n // 100]
                        n %= 100
                        if n > 0:
                            texto += ' e '

                    if n >= 20:
                        texto += dezenas[n // 10]
                        n %= 10
                        if n > 0:
                            texto += ' e ' + unidades[n]
                    elif 10 <= n <= 19:
                        texto += especiais[n]
                    elif n > 0:
                        texto += unidades[n]

                    return texto

                if valor_int == 0:
                    texto_inteiro = 'zero'
                elif valor_int < 1000:
                    texto_inteiro = converter_ate_999(valor_int)
                elif valor_int < 1000000:
                    milhares = valor_int // 1000
                    resto = valor_int % 1000
                    texto_inteiro = converter_ate_999(milhares) + ' mil'
                    if resto > 0:
                        if resto < 100:
                            texto_inteiro += ' e '
                        else:
                            texto_inteiro += ' '
                        texto_inteiro += converter_ate_999(resto)
                else:
                    milhoes = valor_int // 1000000
                    resto = valor_int % 1000000
                    texto_inteiro = converter_ate_999(milhoes) + ' milhões'
                    if resto > 0:
                        if resto < 100:
                            texto_inteiro += ' e '
                        else:
                            texto_inteiro += ' '
                        if resto < 1000:
                            texto_inteiro += converter_ate_999(resto)
                        else:
                            milhares = resto // 1000
                            resto_mil = resto % 1000
                            texto_inteiro += converter_ate_999(milhares) + ' mil'
                            if resto_mil > 0:
                                if resto_mil < 100:
                                    texto_inteiro += ' e '
                                else:
                                    texto_inteiro += ' '
                                texto_inteiro += converter_ate_999(resto_mil)

                texto = texto_inteiro + ' reais'
                if centavos > 0:
                    if centavos == 1:
                        texto += ' e um centavo'
                    else:
                        texto += f' e {converter_ate_999(centavos)} centavos'

                return texto
            except:
                return f"{valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

        def formatar_moeda_br(valor):
            return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

        # PEGAR OS VALORES ORIGINAIS (QUEBRADOS) DAS ÁREAS
        area_app_original = float(dados.get('app_area_original', dados.get('app_area', '0')).replace(',', '.') or '0')
        area_rl_original = float(dados.get('rl_area_original', dados.get('rl_area', '0')).replace(',', '.') or '0')
        area_fora_original = float(
            dados.get('fora_area_original', dados.get('fora_area', '0')).replace(',', '.') or '0')
        area_dano_original = float(
            dados.get('dano_area_original', dados.get('dano_area', '0')).replace(',', '.') or '0')

        # Valores arredondados
        area_app_arredondada = math.ceil(area_app_original) if area_app_original > 0 else 0
        area_rl_arredondada = math.ceil(area_rl_original) if area_rl_original > 0 else 0
        area_fora_arredondada = math.ceil(area_fora_original) if area_fora_original > 0 else 0
        area_dano_arredondada = math.ceil(area_dano_original) if area_dano_original > 0 else 0

        # Calcular hectares completos e fração
        inteiro_app = int(area_app_original)
        fracao_app = area_app_original - inteiro_app

        inteiro_rl = int(area_rl_original)
        fracao_rl = area_rl_original - inteiro_rl

        inteiro_fora = int(area_fora_original)
        fracao_fora = area_fora_original - inteiro_fora

        inteiro_dano = int(area_dano_original)
        fracao_dano = area_dano_original - inteiro_dano

        def formatar_fracao(fracao):
            if fracao <= 0:
                return ""
            return f" + {fracao:.2f} de fração"


        params_app = self.obter_parametros_valoracao_app()

        # Valoração dinâmica: usa o percentual mínimo da faixa/nível efetivamente escolhidos
        acrescimo_min_app = 50000 * (params_app["perc_min"] / 100)
        valor_ha_app = 5000 + acrescimo_min_app
        valor_ai_app = valor_ha_app * area_app_arredondada
        valor_ai_rl = 5000 * area_rl_arredondada
        valor_ai_fora = 1000 * area_fora_arredondada
        valor_ai_dano = 300 * area_dano_arredondada

        total_autos = valor_ai_app + valor_ai_rl + valor_ai_fora + valor_ai_dano
        qte_autos = (1 if area_app_original > 0 else 0) + (1 if area_rl_original > 0 else 0) + (
            1 if area_fora_original > 0 else 0) + (1 if area_dano_original > 0 else 0)

        # Obter números dos autos e embargos
        auto_app = dados.get('app_auto', '') or ''
        auto_rl = dados.get('rl_auto', '') or ''
        auto_fora = dados.get('fora_auto', '') or ''
        auto_dano = dados.get('dano_auto', '') or ''

        embargo_app = dados.get('app_embargo', '') or ''
        embargo_rl = dados.get('rl_embargo', '') or ''
        embargo_fora = dados.get('fora_embargo', '') or ''

        intervalo_supressao = dados.get('intervalo_supressao', 'período não informado')

        # CONSTRUIR O RELATÓRIO - APENAS AS ÁREAS SELECIONADAS
        relatorio = f"""Em cumprimento a Ordem de Serviço SEMAD/GO de nº {dados['os']} a Gerência de Geoprocessamento e Sensoriamento Remoto (GEGEO) realizou análise geoespacial em torno do(s) alerta(s) de desmatamento de número(s) {dados['alertas']} do Mapbiomas, visando monitoramento e tomada de medidas cabíveis sobre o desmatamento no Estado de Goiás.

    No dia {dados['data']}, foi realizada fiscalização remota do imóvel rural denominado {dados['imovel']} com número de CAR: {dados['car']}, coordenadas SIRGAS 2000 ({dados['coordenadas']}), município de {dados['municipio']}- GO e por meio de análise geoespacial foram identificadas supressões, passíveis de autuação, nas seguintes áreas:
        • Área de Preservação Permanente (APP): {area_app_original:.2f} hectares;
        • Área de Reserva Legal (RL): {area_rl_original:.2f} hectares;
        • Área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP): {area_fora_original:.2f} hectares;
        • Dano à vegetação nativa (corte de árvores isoladas): {area_dano_original:.2f} hectares.

    As supressões ocorreram {intervalo_supressao}, e a descrição detalhada de cada polígono encontra-se especificada no Mapa Geral e no Mapa de Análise Temporal anexados a este relatório. 
    No momento da fiscalização, foi constatado que as supressões que deram origem as infrações estão desprovidas de Licenciamento Ambiental e que o Imóvel rural cadastrado no Sistema Nacional de Cadastro Ambiental Rural - SICAR encontra-se em nome de {dados['proprietario']}, CPF: {dados['cpf']}, o que ensejou a tomada das medidas administrativas cabíveis. Diante de todos esses elementos, originou-se a responsabilização e aplicação dos Autos de Infração e Termos de Embargo instrumentos estes vinculados a este relatório. Tais medidas foram adotadas visando cessar o dano causado ao meio ambiente em que a reparação e/ou compensação pelo dano será estabelecida por meio de medidas técnicas e ambientais no âmbito da licença ambiental do empreendimento ou de termo de compromisso específico.

    """

        # ==========================================================
        # ÁREA DE PRESERVAÇÃO PERMANENTE (APP)
        # ==========================================================
        if area_app_original > 0:
            relatorio += f"""
    {'=' * 60}

    ÁREA DE PRESERVAÇÃO PERMANENTE

    1) - Auto de Infração nº {auto_app} - Estabelecido pelo artigo 43 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 5.000,00 (cinco mil reais) a R$ 50.000,00 (cinquenta mil reais) por hectare ou fração (multa aberta) e para fixação do valor do auto, referente a supressão de {area_app_original:.2f} hectares em Área de Preservação Permanente (APP), foram utilizados os critérios definidos na Orientação Normativa SEMAD Nº 1/2024 considerando o nível de gravidade da infração e identificação da capacidade econômica do infrator da seguinte maneira:

    GRAVIDADE
    I - Motivo da Infração: pontuação = 10 (intencional, pois realizou a supressão ciente que não tinha autorização para tal);
    II - Consequência para o meio ambiente: pontuação = 30 (moderada, pois supressão de uma Área de Preservação Permanente - APP compromete a função ambiental de preservar os recursos hídricos);
    III - Consequência para a saúde pública ou para socioeconomia da área de abrangência do fato: pontuação = 0 (não foram constatadas evidências de consequência para a saúde pública ou para a socioeconomia da área de abrangências do fato).
    Somatório das etapas = 40 pontos, com classificação Nível B.

    SITUAÇÃO ECONÔMICA
    Não foi possível identificar a capacidade econômica do autuado pela ausência de documentos ou informações, neste caso optou-se pela {params_app["faixa"]} ({params_app["tipo_infrator"]} - Receita Bruta mensal {params_app["descricao_faixa"]}, {params_app["nivel"]}, mínimo + {params_app["perc_min"]:.3f}% até {params_app["perc_max"]:.3f}% do teto), no entanto, no momento da audiência de autocomposição poderá ser reclassificado conforme a capacidade econômica, mediante comprovação por documentos, conforme § 2º do Art. 8 da ON - 01/2024.
    
    Diante do exposto chegou-se a seguinte VALORAÇÃO: (R$ 5.000,00 + ({params_app["perc_min"]:.3f}% X R$ 50.000,00)) x {area_app_original:.2f} hectares ({inteiro_app} hectares completos{formatar_fracao(fracao_app)}), ou seja, R$ {valor_ha_app:.2f} x {area_app_arredondada}, totalizando o valor de {formatar_moeda_br(valor_ai_app)} ({numero_extenso(valor_ai_app)}), ficando esta área embargada pelo Termo de Embargo de nº {embargo_app}.
    """

        # ==========================================================
        # RESERVA LEGAL (RL)
        # ==========================================================
        if area_rl_original > 0:
            relatorio += f"""
    {'=' * 60}

    RESERVA LEGAL

    2) - Auto de Infração nº {auto_rl} - Estabelecido pelo artigo 51 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 5.000,00 (cinco mil reais) por hectare ou fração (multa fechada), e pela supressão de {area_rl_original:.2f} hectares em área de Reserva Legal (RL) o auto foi valorado da seguinte forma: R$ 5.000,00 x {area_rl_original:.2f} hectares ({inteiro_rl} hectares completos{formatar_fracao(fracao_rl)}), ou seja, R$ 5.000,00 x {area_rl_arredondada}, totalizando o valor de {formatar_moeda_br(valor_ai_rl)} ({numero_extenso(valor_ai_rl)}), ficando esta área embargada pelo Termo de Embargo de nº {embargo_rl}.
    """

        # ==========================================================
        # FORA DE APP E RL
        # ==========================================================
        if area_fora_original > 0:
            relatorio += f"""
    {'=' * 60}

    FORA DE APP E RL

    3) - Auto de Infração nº {auto_fora} - Estabelecido pelo artigo 52 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 1.000,00 (um mil reais) por hectare ou fração (multa fechada), e pela supressão de {area_fora_original:.2f} hectares em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP) o auto foi valorado da seguinte forma: R$ 1.000,00 x {area_fora_original:.2f} hectares ({inteiro_fora} hectares completos{formatar_fracao(fracao_fora)}), ou seja, R$ 1.000,00 x {area_fora_arredondada}, totalizando o valor de {formatar_moeda_br(valor_ai_fora)} ({numero_extenso(valor_ai_fora)}), ficando esta área embargada pelo Termo de Embargo de nº {embargo_fora}.
    """

        # ==========================================================
        # DANO AMBIENTAL
        # ==========================================================
        if area_dano_original > 0:
            relatorio += f"""
    {'=' * 60}

    DANO AMBIENTAL

    4) - Auto de Infração nº {auto_dano} – Estabelecido pelo artigo 53 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 300,00 (trezentos reais) por hectare ou fração, ou por unidade, estéreo, quilo, mdc ou metro cúbico (multa fechada), e pelo dano (corte de árvores isoladas) de {area_dano_original:.2f} hectares em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP) o auto foi valorado da seguinte forma: R$ 300,00 x {area_dano_original:.2f} hectares ({inteiro_dano} hectares completos{formatar_fracao(fracao_dano)}), ou seja, R$ 300,00 x {area_dano_arredondada}, totalizando o valor de {formatar_moeda_br(valor_ai_dano)} ({numero_extenso(valor_ai_dano)}). Esta área não fica embargada.
    """

        # ==========================================================
        # SOMA TOTAL DOS AUTOS
        # ==========================================================
        relatorio += f"""
    {'=' * 60}

    SOMA TOTAL DOS AUTOS

    A soma dos {qte_autos} autos totaliza o valor de {formatar_moeda_br(total_autos)} ({numero_extenso(total_autos)}).

    """

        # ==========================================================
        # APLICAÇÃO DO ART. 93 - UNIDADE DE CONSERVAÇÃO (UC)
        # ==========================================================
        # VERIFICA SE O CHECKBOX UC ESTÁ MARCADO
        if self.is_uc_marcado():
            relatorio += f"""
    {'=' * 60}
    📜 APLICAÇÃO DO ART. 93 - UNIDADE DE CONSERVAÇÃO (UC)

    Conforme estabelecido no Art. 93 do Decreto Federal nº 6.514/2008:

    "Art. 93. As infrações administrativas praticadas em Unidade de Conservação 
    ou na sua Zona de Amortecimento terão os valores das multas aplicadas em dobro."

    Portanto, por se tratar de infração ocorrida em Unidade de Conservação 
    ou em sua Zona de Amortecimento, os valores das multas foram aplicados 
    em DOBRO, conforme determinado pela legislação.

    O valor total dos autos ({formatar_moeda_br(total_autos)}) já contempla 
    a majoração prevista no Art. 93.

    """

        # ==========================================================
        # FINALIZAÇÃO
        # ==========================================================
        relatorio += f"""
    {'-' * 60}
    {dados['municipio']}, {dados['data']}
    _____________________________________
    Assinatura do Fiscal
    """

        return relatorio

    def gerar_relatorio_barramento(self, dados_especificos):
        """Gera o relatório de barramento completo com dados integrados das abas"""
        dados = self.obter_dados_comuns()
        dados.update(dados_especificos)

        # ==========================================================
        # BUSCA DADOS DO ART. 66 E ART. 43 AUTOMATICAMENTE
        # ==========================================================
        import re

        # 1. BUSCAR AUTO E EMBARGO DO ART. 66
        auto_66 = ""
        embargo_66 = ""
        if "art66" in self.parent.artigos_inputs:
            auto_widget = self.parent.artigos_inputs["art66"].get("auto")
            if auto_widget and hasattr(auto_widget, 'text'):
                auto_66 = auto_widget.text().strip()

            embargo_widget = self.parent.artigos_inputs["art66"].get("embargo")
            if embargo_widget and hasattr(embargo_widget, 'text'):
                embargo_66 = embargo_widget.text().strip()

        # Se não encontrou, usa os dados da interface
        if not auto_66:
            auto_66 = dados.get('auto_artigo_66', 'NÃO INFORMADO')
        if not embargo_66:
            embargo_66 = dados.get('embargo_barramento', 'NÃO INFORMADO')

        # 2. BUSCAR VALOR DO ART. 66
        valor_66 = 3000.00
        if "art66" in self.parent.artigos_inputs:
            valor_widget = self.parent.artigos_inputs["art66"].get("valor")
            if valor_widget and hasattr(valor_widget, 'text'):
                valor_text = valor_widget.text().strip()
                if valor_text:
                    numeros = re.findall(r'[\d\.,]+', valor_text)
                    if numeros:
                        try:
                            valor_str = numeros[0].replace('.', '').replace(',', '.')
                            valor_66 = float(valor_str)
                        except:
                            pass

        # 3. BUSCAR AUTO E EMBARGO DO ART. 43
        auto_43 = ""
        embargo_43 = ""
        if "art43" in self.parent.artigos_inputs:
            auto_widget = self.parent.artigos_inputs["art43"].get("auto")
            if auto_widget and hasattr(auto_widget, 'text'):
                auto_43 = auto_widget.text().strip()

            embargo_widget = self.parent.artigos_inputs["art43"].get("embargo")
            if embargo_widget and hasattr(embargo_widget, 'text'):
                embargo_43 = embargo_widget.text().strip()

        # Se não encontrou, usa os dados da interface
        if not auto_43:
            auto_43 = dados.get('auto_artigo_43', 'NÃO INFORMADO')
        if not embargo_43:
            embargo_43 = dados.get('embargo_barramento', 'NÃO INFORMADO')

        # 4. BUSCAR VALOR DO ART. 43
        valor_43 = 5012.50
        if "art43" in self.parent.artigos_inputs:
            valor_widget = self.parent.artigos_inputs["art43"].get("valor")
            if valor_widget and hasattr(valor_widget, 'text'):
                valor_text = valor_widget.text().strip()
                if valor_text:
                    numeros = re.findall(r'[\d\.,]+', valor_text)
                    if numeros:
                        try:
                            valor_str = numeros[0].replace('.', '').replace(',', '.')
                            valor_43 = float(valor_str)
                        except:
                            pass

        # 5. VALOR RECURSOS HÍDRICOS (vem da interface)
        try:
            valor_recursos = dados.get('valor_recursos_hidricos', '902.25')
            if isinstance(valor_recursos, str):
                valor_recursos = float(valor_recursos.replace(',', '.'))
            elif not valor_recursos:
                valor_recursos = 902.25
        except:
            valor_recursos = 902.25

        # 6. ÁREA DE SUPRESSÃO
        try:
            area_supressao = dados.get('area_supressao', '0')
            if isinstance(area_supressao, str):
                area_supressao = float(area_supressao.replace(',', '.'))
            elif not area_supressao:
                area_supressao = 0
        except:
            area_supressao = 0

        # Calcula área arredondada
        area_arredondada = math.ceil(area_supressao) if area_supressao > 0 else 1

        # 7. INTERVALO DA SUPRESSÃO (vem dos Dados Gerais)
        intervalo_supressao = dados.get('intervalo_supressao', 'período não informado')

        # 8. DADOS DA ÁREA DO BARRAMENTO
        try:
            area_barramento = dados.get('area_barramento', '0')
            if isinstance(area_barramento, str):
                area_barramento = float(area_barramento.replace(',', '.'))
            elif not area_barramento:
                area_barramento = 0
        except:
            area_barramento = 0

        # 9. CALCULA TOTAL DOS AUTOS
        total_autos = valor_66 + valor_43 + valor_recursos

        # 10. OUTROS DADOS
        data_ocorrencia = dados.get('data_ocorrencia', datetime.now().strftime("%d/%m/%Y"))
        data_sobrevoo = dados.get('data_sobrevoo', 'NÃO INFORMADO')
        mapa_geral = dados.get('mapa_geral', 'NÃO INFORMADO')
        mapa_temporal = dados.get('mapa_temporal', 'NÃO INFORMADO')
        manifestacao = dados.get('manifestacao', 'NÃO INFORMADO')
        fato_denunciado = dados.get('fato_denunciado', 'NÃO INFORMADO')
        texto_ina_pr = dados.get('texto_ina_pr', 'NÃO INFORMADO')

        # 11. DADOS GERAIS
        processo = dados.get('processo', 'NÃO INFORMADO')
        os_num = dados.get('os', 'NÃO INFORMADO')
        imovel = dados.get('imovel', 'NÃO INFORMADO')
        municipio = dados.get('municipio', 'NÃO INFORMADO')
        uf = dados.get('uf', 'GO')
        car = dados.get('car', 'NÃO INFORMADO')
        coordenadas = dados.get('coordenadas', 'NÃO INFORMADO')
        proprietario = dados.get('proprietario', 'NÃO INFORMADO')
        cpf = dados.get('cpf', 'NÃO INFORMADO')

        # ==========================================================
        # CONSTRUÇÃO DO RELATÓRIO
        # ==========================================================
        relatorio = f"""RELATÓRIO DE FISCALIZAÇÃO - BARRAMENTO

    Em cumprimento a Ordem de Serviço SEMAD/GO de nº {os_num}, no dia {data_ocorrencia}, esta equipe de fiscalização deslocou-se até o município de {municipio} para averiguar as informações prestadas na {manifestacao}, cuja descrição: {fato_denunciado}

    Conforme informado no INÃ PR:
    "{texto_ina_pr}"

    Para cumprir com os objetivos estabelecidos, foram empregadas diversas metodologias de fiscalização, incluindo análise de imagens de satélite, observação in loco e documental, visando uma apuração detalhada e precisa da situação.

    Tipo da Ação:
    ( ) Análise/Fiscalização Processual
    ( ) Fiscalização (remota) – §4° do Art.36 da Lei Estadual 20.694/2019
    ( ) Fiscalização in loco
    ( ) Fiscalização em atividade sem licença.

    Motivação:
    ( ) Acompanhamento
    ( ) Análise Complementar (juntada de documento após fiscalização)
    ( ) Análise de Atendimento de Notificação
    ( ) Denúncia/Ouvidoria/INÃ: {manifestacao}
    ( ) Análise de Processos solicitados por órgãos externos – Processo:
    ( ) LAI - Lei de Acesso à Informação – Processo:

    Localização da Atividade:
    ( ) Zona Urbana: {imovel}, no município {municipio}, no entorno da Coordenada Geográfica SIRGAS 2000: {coordenadas}
    ( ) Zona Rural: {imovel}, no município {municipio}, no entorno da Coordenada Geográfica SIRGAS 2000: {coordenadas}

    Descrição da atividade fiscalizatória
    HISTÓRICO PROCESSUAL:
    SGA: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    IPÊ: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    SEI: ( ) Nada consta ( ) Não se aplica ( X ) Processo(s)/Cadastro(s): {processo}
    INÃ: ( ) Nada consta ( ) Não se aplica ( X ) Processo(s)/Cadastro(s): {processo}
    WebOutorga / Veredas: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    SIGA: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):
    SICAR: ( ) Nada consta ( ) Não se aplica ( X ) Processo(s)/Cadastro(s): {car}
    SEISB: ( ) Nada consta ( ) Não se aplica ( ) Processo(s)/Cadastro(s):

    CONSTATAÇÕES
    Constatações Técnicas da Fiscalização
    ( ) A equipe foi recebida por: {proprietario};
    ( ) Ninguém foi encontrado no local.

    Observações:
    Área localizada na zona rural do município de {municipio};
    Na propriedade, foi constatada a realização de obras em uma represa localizada no entorno da Coordenada Geográfica SIRGAS 2000: {coordenadas}. O {proprietario} declarou tratar-se de uma represa antiga e que, juntamente com o proprietário da área confrontante (lado oposto), realizou apenas a revitalização da área, há aproximadamente seis meses.

    No momento da fiscalização, constatou-se que o reservatório foi objeto de intervenções recentes, sendo possível observar sinais de movimentação de terra e taludes ainda sem cobertura vegetal, que no maciço do talude houve plantio recente de grama. O reservatório possui 02 (dois) extravasores nas ombreiras direita e esquerda e descarga de fundo, por meio de 02 (dois) tubos de PVC no centro do barramento. No momento da fiscalização, observou-se o vertimento de água pela descarga de fundo e por ambas as ombreiras.

    Na mesma data, tendo em vista que a denúncia tratava-se de várias intervenções ao longo do curso hídrico, após fiscalização no proprietário do outro lado da barragem, o mesmo informou que nessa localidade a obra foi efetuada pelo {proprietario}, o qual não se responsabilizou no momento da fiscalização. Mas foi observado in loco que o {proprietario} é o principal usuário, sendo percebido uma embarcação e algumas tambores para ração, os quais são utilizados para alimentar os peixes colocados no barramento.

    Foram realizadas a coleta de dados in loco e, posteriormente, a consolidação e análise das informações obtidas.

    Foi realizado sobrevoo da área por meio de Aeronave Remotamente Pilotada - RPA, modelo DJI Mavic 3, no dia {data_sobrevoo}, com intuito de coletar imagens aéreas das áreas em questão, possibilitando a confecção de ortofoto das áreas de interesse por meio do software WebODM e a realização de análises de geoprocessamento por meio do software QGIS. Foi elaborado mapa digital de situação contendo as áreas de interesse.

    Com objetivo de subsidiar a atividade de fiscalização ambiental foi elaborado o seguinte produto cartográfico:
    - Mapa_Geral - {mapa_geral}
    - Mapa_Temporal - {mapa_temporal}

    Para a análise da área foram utilizadas imagens de janeiro, maio, e junho de 2025, da constelação Planet com resolução espacial de 4,77 m, disponibilizadas pela Iniciativa Internacional sobre Clima e Florestas da Noruega-NICFI (https://www.planet.com/nicfi/). Foram utilizados dados disponibilizados no Sistema de Informações Geográficas do Estado de Goiás – SIGA. Todas as informações geoespaciais estão georreferenciadas ao Sistema de Coordenadas UTM 22S, Datum SIRGAS2000, utilizando o software livre de Geoprocessamento QGIS.

    Constatado:
    1 reservatório/barramento:
    - Lâmina d'água: {area_barramento:.2f} hectares;
    - Supressão/Dano: {area_supressao:.2f} hectares {intervalo_supressao}.

    Diante do exposto, conclui-se que a operação do reservatório em questão encontra-se desprovida de licença ambiental válida, bem como inexistente outorga de direito de uso de recursos hídricos para a acumulação hídrica, em desacordo com a legislação ambiental vigente. Ressalta-se que os danos à vegetação nativa constatados na área decorrem diretamente da implantação do barramento. A vegetação existente foi suprimida durante a execução da obra e, posteriormente, danificada pelo acúmulo permanente de água, o que ocasionou a morte da cobertura vegetal remanescente. Ademais, a formação do reservatório alterou as condições naturais do ambiente, impedindo a regeneração natural da vegetação nas áreas atingidas, especialmente na Área de Preservação Permanente (APP).

    Em razão das irregularidades constatadas, foram adotadas as medidas legais cabíveis, nos termos da normativa aplicável, visando à regularização da atividade e à proteção dos recursos ambientais. O presente relatório consolida os levantamentos e análises realizadas, subsidiando os encaminhamentos administrativos pertinentes no âmbito desta Secretaria.

    Sendo assim, foram lavrados em nome de {proprietario}, CPF: {cpf}, os Autos de Infração nº {auto_43} e {auto_66} e Termos de Embargo {embargo_43} e {embargo_66}, e devidamente enviados por meio de carta registrada com aviso de recebimento, para o endereço de correspondência informado pelo {proprietario} no momento da fiscalização.

    LEGISLAÇÃO PERTINENTE
    - Decreto Federal nº 6.514/2008;
    - Decreto Estadual nº 9.710/2020;
    - Decreto Estadual nº 10.371/2023;
    - Lei Federal nº 9.605/1998;
    - Lei Estadual nº 20.694/2019;
    - Lei Estadual nº 18.102/2013;
    - Lei Estadual nº 13.123/1997.

    INFORMAÇÕES COMPLEMENTARES:

    DECRETO Nº 10.371, DE 20 DE DEZEMBRO DE 2023
    Altera o Decreto nº 9.710, de 3 de setembro de 2020, que regulamenta, no âmbito do Poder Executivo estadual, a Lei nº 20.694, de 26 de dezembro de 2019, que dispõe sobre as normas gerais para o licenciamento ambiental no Estado de Goiás e dá outras providências.

    ANEXO ÚNICO (DECRETO Nº 9.710, DE 3 DE SETEMBRO DE 2020)
    DIVISÃO "F": OBRAS CIVIS
    Grupo F2: barragens, diques e canais.
    F2.1 - Reservatórios e diques para captação de água de chuva ou derivada, fora de APP e leito de rio perene ou intermitente - Lâmina de água do reservatório (ha) - Micro ≥ 1 < 5
    F2.2 - Reservatórios/barragens e diques em curso de água para abastecimento humano, dessedentação animal, irrigação, fins paisagísticos* e composição urbana, lazer, turismo e aquicultura sem remoção de pessoas. * para fins paisagísticos e composição urbana, lazer ou turismo, somente com decreto do Chefe do Poder Executivo estadual ou federal; e ** as barragens instaladas depois de 27 de dezembro de 2019 com área do reservatório menor do que 1,2 ha e para os fins descritos acima deverão ser enquadradas na tipologia F2.6 - Lâmina de água do reservatório (ha) Micro ≥ 1,2** < 5
    F2.6 - Reservatórios/barragens e diques em curso de água com lâmina d'água entre 0,1 e 1,2 hectare para abastecimento humano, dessedentação animal, irrigação, fins paisagísticos* e composição urbana, lazer, turismo e aquicultura sem remoção de pessoas. * para fins paisagísticos e outros usos não previstos em lei, é necessário o decreto de utilidade pública - Lâmina de água do reservatório (ha) - Micro ≥ 0,1 < 1,2.

    LEI COMPLEMENTAR Nº 140, DE 8 DE DEZEMBRO DE 2011
    No artigo 3º da Resolução CEMAm 259/2024, especifica quais os parâmetros e requisitos o município deverá atender para o credenciamento para licenciar as atividades que estão definidas no anexo único da mesma resolução.
    {municipio} - Res. CEMAm n°174, de 18 de Outubro de 2022 - Nível 2.

    1. Que atividade(s) é(são) realizada(s) no local?
    Barramento em área de APP e acumulação de água sem autorização do órgão ambiental competente;

    2. A(s) atividade(s) é(são) utilizadora(s) de recursos ambientais, consideradas efetiva ou potencialmente poluidoras, ou capazes de, sob qualquer forma, causar degradação ambiental?
    Sim;

    3. A(s) atividade(s) é(são) licenciada(s)/autorizada(s)?
    Não;

    4. Quem é(são) o(s) responsável(is) pela(s) atividade(s)?
    {proprietario}, CPF: {cpf};

    5. Qual o endereço e coordenadas geográficas do local onde é(são) realizada(s) a(s) atividade(s)?
    {imovel}, no município de {municipio}, no entorno da Coordenada Geográfica SIRGAS 2000: {coordenadas};

    6. Quem é(são) o(s) proprietário(s) da(s) área(s) onde é(são realizadas(s) a(s) atividade(s)?
    {proprietario}, CPF: {cpf};

    7. A(s) atividade(s) é(são) desenvolvida(s) em áreas protegidas (APP, Reserva Legal e/ou Unidade de Conservação)?
    Sim;

    8. A(s) atividade(s) causou(ram) ou está(ão) causando degradação ambiental? Se sim, em que consiste a degradação ambiental?
    Sim. Consiste na operação de barragem em área de APP, sem autorização do órgão ambiental competente;

    9. Que medidas foram ou devem ser adotadas para cessar a degradação ambiental?
    Foi realizado o embargo da área e das atividades nelas realizadas;

    10. A degradação ambiental comporta recuperação?
    Sim, a degradação ambiental pode ser passível de recuperação, embora a viabilidade e o tempo necessário para a recuperação possam variar dependendo da extensão e gravidade da degradação, bem como das características do ecossistema afetado;

    11. Que medidas foram ou devem ser adotadas para promover a recuperação ambiental?
    As medidas técnicas e ambientais serão estabelecidas no âmbito da licença ambiental do empreendimento ou do termo de compromisso específico;

    DA VALORAÇÃO

    Art. 66 - Auto de Infração nº {auto_66}

    O artigo 66 do Decreto nº 6.514/2008 estabelece a valoração de Multa de R$ 500,00 (quinhentos reais) a R$ 10.000.000,00 (dez milhões de reais):
    Art. 66. Construir, reformar, ampliar, instalar ou fazer funcionar estabelecimentos, atividades, obras ou serviços utilizadores de recursos ambientais, considerados efetiva ou potencialmente poluidores, sem licença ou autorização dos órgãos ambientais competentes, em desacordo com a licença obtida ou contrariando as normas legais e regulamentos pertinentes:
    Multa de R$ 500,00 (quinhentos reais) a R$ 10.000.000,00 (dez milhões de reais).

    Para fixação do valor referente a infração foram utilizados os critérios definidos na ORIENTAÇÃO NORMATIVA SEMAD Nº 01/2024, considerando o nível de gravidade da infração e identificação da capacidade econômica da seguinte maneira:
    - Motivo da Infração: Obtenção de vantagem pecuniária (15)
    - Consequência para o meio ambiente: Moderada (30)
    - Consequência para a saúde pública ou para a socioeconomia da área de abrangência do fato: Não Houve (0)
    Somatório dos valores desta etapa: (45) - Nível B
    Situação econômica: Não foi possível aferir a situação econômica do infrator. Por conta disso, foi utilizado Situação econômica - Receita Mensal - Pessoa Física - Receita bruta: mensal de até 1 salário mínimo (Faixa A) = Nível B: Mínimo + 0,025% até 0,8% do teto.

    VALORAÇÃO: R$ 500,00 + (0,025% X R$ 10.000.000,00) = {self._formatar_moeda_br(valor_66)} ({self._numero_extenso(valor_66)})

    Considerando-se que foi utilizado a classe de menor valor para a situação econômica do infrator, que pode não ser condizente com a realidade, sugere-se na Autocomposição/julgamento, a aferição e reenquadramento se for o caso conforme Art. 8º, §2º ON 01/2024 SEMAD.
    Art. 8º Em se tratando de pessoa física, a situação econômica do infrator será determinada pelos critérios estabelecidos no
    Quadro 3 do Anexo único, mediante a classificação em faixas definidas conforme receita bruta anual do infrator, assim estabelecidas:
    I - receita bruta mensal de até 1 salário mínimo;
    II - receita bruta mensal, superior ao limite anterior até 3 salários mínimos;
    III - receita bruta mensal, superior ao limite anterior até 10 salários mínimos;
    IV - receita bruta mensal, superior ao limite anterior até 30 salários mínimos;
    V - receita bruta mensal, superior ao limite anterior até 45 salários mínimos; e
    VI - receita bruta mensal, superior ao limite anterior.
    § 1º Em se tratando de pessoa física serão considerados os rendimentos indicados em qualquer documento válido para comprovação de renda.
    § 2º A autoridade julgadora competente bem como os facilitadores em sede de audiências de autocomposição deverão rever o enquadramento do infrator quanto a sua situação econômica, caso conste no relatório de fiscalização que esta não tenha sido possível aferir.

    Lei Estadual 13.123, DE 16 DE JULHO DE 1997.
    SEÇÃO II DAS INFRAÇÕES E PENALIDADES
    Art. 13. Constitui infração às normas de utilização de recursos hídricos superficiais e subterrâneos:
    I - derivar ou utilizar dos recursos hídricos para qualquer finalidade, sem a respectiva outorga de direito de uso;
    II - iniciar a implantação ou implantar empreendimento relacionado com a derivação ou utilização de recursos hídricos, superficiais e/ou subterrâneos, que implique alterações no regime, quantidade e qualidade dos mesmos, sem autorização dos órgãos ou entidades competentes;
    III - deixar expirar o prazo de validade das outorgas sem solicitar a devida prorrogação ou revalidação;
    IV - utilizar-se dos recursos hídricos ou executar obras ou serviços relacionados com os mesmos em desacordo com as condições estabelecidas na outorga;
    Art. 14. Por infração a qualquer disposição legal ou regulamentar referente a execução de obras e serviços hidráulicos, derivação e utilização de recursos hídricos de domínio ou administração do Estado de Goiás, ou pelo não atendimento das solicitações feitas, o infrator, a critério da autoridade competente, ficará sujeito às seguintes penalidades, independentemente da sua ordem de enumeração:
    II - multa, simples ou diária, proporcional à gravidade da infração, de R$ 90,00 (noventa reais) a R$ 90.000,00 (noventa mil reais), corrigidos pela UFIR;
    Art. 15. As infrações às disposições desta lei às normas dela decorrentes serão, a critério da autoridade impositora, classificadas em leves, graves, gravíssimas, levando em conta:
    I - as circunstâncias atenuantes e agravantes;
    II - os antecedentes do infrator.
    § 1o As multas simples ou diárias, a critério da autoridade aplicadora, ficam estabelecidas dentro das seguintes faixas:
    b) acima de R$ 900,00 (novecentos reais) até 9.000,00 (nove mil reais), nas infrações graves;

    Resolução CERHi nº 66, de 26 de janeiro de 2024, artigo 3°, inciso V:
    Art. 3°. Estão sujeitos à outorga:
    V – as acumulações de água em corpos hídricos;
    Para fixação do valor referente a infração foi utilizado os critérios definidos nos quadros 1 e 3 da ORIENTAÇÃO NORMATIVA SEMAD Nº 1/2024 - GAB- 06281 e por não contemplar as classificações em infrações leves, graves e gravíssimas referente ao funcionamento/operação de barragem sem a respectiva outorga de direito de uso, foi utilizado por analogia a RESOLUÇÃO Nº 24, DE 04 DE MAIO DE 2020 da Agência Nacional das Águas - ANA, considerando o inciso I, do Artigo 18, o nível de gravidade da infração e identificação da capacidade econômica da seguinte maneira:
    Art. 18. São consideradas infrações graves:
    I – derivar ou utilizar recursos hídricos para qualquer finalidade, sem a respectiva outorga de direito de uso;

    Para fixação do valor referente a infração foram utilizados os critérios definidos na ORIENTAÇÃO NORMATIVA SEMAD Nº 01/2024 (Define parâmetros para a fixação das multas abertas, para a aplicação de sanções e medidas administrativas cautelares no âmbito da apuração de infrações ambientais, bem como critérios para o agravamento e a atenuação das sanções administrativas decorrentes de infrações ambientais), considerando o nível de gravidade da infração e identificação da capacidade econômica. Gravidade da Infração:
    - Motivo da Infração: Obtenção de vantagem pecuniária (15)
    - Consequência para o meio ambiente: Fraca (20)
    - Consequência para a saúde pública ou para a socioeconomia da área de abrangência do fato: Não Houve (0)
    Somatório dos valores desta etapa: (35) - Nível B
    Situação econômica: Não foi possível aferir a situação econômica do infrator. Por conta disso, foi utilizado Situação econômica - Receita Mensal - Pessoa Física - Receita bruta: mensal de até 1 salário mínimo (Faixa A) = Nível B: Mínimo + 0,025% até 0,8% do teto.

    VALORAÇÃO: (R$ 900,00 + (0,025% X R$ 9.000,00)), totalizando o valor de {self._formatar_moeda_br(valor_recursos)} ({self._numero_extenso(valor_recursos)}).

    Considerando-se que foi utilizado a classe de menor valor para a situação econômica do infrator, que pode não ser condizente com a realidade, sugere-se na Autocomposição/julgamento, a aferição e reenquadramento se for o caso conforme Art. 8º, §2º ON 01/2024 SEMAD.

    Art. 43 - Auto de Infração nº {auto_43}

    O artigo 43 do Decreto nº 6.514/2008 estabelece a valoração de Multa de R$ 5.000,00 (cinco mil reais) a R$ 50.000,00 (cinquenta mil reais), por hectare ou fração:
    Art. 43. Destruir ou danificar florestas ou demais formas de vegetação natural ou utilizá-las com infringência das normas de proteção em área considerada de preservação permanente, sem autorização do órgão competente, quando exigível, ou em desacordo com a obtida: 
    Para fixação do valor referente a infração foram utilizados os critérios definidos na ORIENTAÇÃO NORMATIVA SEMAD Nº 01/2024 (Define parâmetros para a fixação das multas abertas, para a aplicação de sanções e medidas administrativas cautelares no âmbito da apuração de infrações ambientais, bem como critérios para o agravamento e a atenuação das sanções administrativas decorrentes de infrações ambientais), considerando o nível de gravidade da infração e identificação da capacidade econômica da seguinte maneira:
    - Motivo da Infração: Obtenção de vantagem pecuniária (15)
    - Consequência para o meio ambiente: Moderada (30)
    - Consequência para a saúde pública ou para a socioeconomia da área de abrangência do fato: Não Houve (0)
    Somatório dos valores desta etapa: (45) - Nível B
    Situação econômica: Não foi possível aferir a situação econômica do infrator. Por conta disso, foi utilizado Situação econômica - Receita Mensal - Pessoa Física - Receita bruta: mensal de até 1 salário mínimo (Faixa A) = Nível B: Mínimo + 0,025% até 0,8% do teto.

    R$ 5.000,00 + (0,025% X R$ 50.000,00) = {self._formatar_moeda_br(valor_43)} ({self._numero_extenso(valor_43)})

    Valoração: {area_supressao:.2f} hectares x {self._formatar_moeda_br(valor_43)} = {self._formatar_moeda_br(valor_43 * area_arredondada)} ({self._numero_extenso(valor_43 * area_arredondada)})

    {'=' * 60}
    SOMA TOTAL DOS AUTOS

    A soma dos autos totaliza o valor de {self._formatar_moeda_br(total_autos)} ({self._numero_extenso(total_autos)}).

    Considerando-se que foi utilizado a classe de menor valor para a situação econômica do infrator, que pode não ser condizente com a realidade, sugere-se na Autocomposição/julgamento, a aferição e reenquadramento se for o caso conforme Art. 8º, §2º ON 01/2024 SEMAD.

    {'=' * 60}
    {municipio}, {data_ocorrencia}
    _____________________________________
    Assinatura do Fiscal
    """
        return relatorio

    def gerar_relatorio_supressao(self, dados):
        """Gera o conteúdo do Relatório INÃ - Supressão"""

        def formatar_moeda_br(valor):
            return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

        def numero_extenso(valor):
            try:
                valor_int = int(valor)
                centavos = int(round((valor - valor_int) * 100))

                unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
                especiais = {10: 'dez', 11: 'onze', 12: 'doze', 13: 'treze', 14: 'quatorze',
                             15: 'quinze', 16: 'dezesseis', 17: 'dezessete', 18: 'dezoito', 19: 'dezenove'}
                dezenas = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta',
                           'sessenta', 'setenta', 'oitenta', 'noventa']
                centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos',
                            'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']

                def converter_ate_999(n):
                    if n == 0:
                        return ''
                    if n == 100:
                        return 'cem'
                    texto = ''
                    if n >= 100:
                        texto += centenas[n // 100]
                        n %= 100
                        if n > 0:
                            texto += ' e '
                    if n >= 20:
                        texto += dezenas[n // 10]
                        n %= 10
                        if n > 0:
                            texto += ' e ' + unidades[n]
                    elif 10 <= n <= 19:
                        texto += especiais[n]
                    elif n > 0:
                        texto += unidades[n]
                    return texto

                if valor_int == 0:
                    texto_inteiro = 'zero'
                elif valor_int < 1000:
                    texto_inteiro = converter_ate_999(valor_int)
                elif valor_int < 1000000:
                    milhares = valor_int // 1000
                    resto = valor_int % 1000
                    texto_inteiro = converter_ate_999(milhares) + ' mil'
                    if resto > 0:
                        if resto < 100:
                            texto_inteiro += ' e '
                        else:
                            texto_inteiro += ' '
                        texto_inteiro += converter_ate_999(resto)
                else:
                    milhoes = valor_int // 1000000
                    resto = valor_int % 1000000
                    texto_inteiro = converter_ate_999(milhoes) + ' milhões'
                    if resto > 0:
                        if resto < 100:
                            texto_inteiro += ' e '
                        else:
                            texto_inteiro += ' '
                        if resto < 1000:
                            texto_inteiro += converter_ate_999(resto)
                        else:
                            milhares = resto // 1000
                            resto_mil = resto % 1000
                            texto_inteiro += converter_ate_999(milhares) + ' mil'
                            if resto_mil > 0:
                                if resto_mil < 100:
                                    texto_inteiro += ' e '
                                else:
                                    texto_inteiro += ' '
                                texto_inteiro += converter_ate_999(resto_mil)

                texto = texto_inteiro + ' reais'
                if centavos > 0:
                    if centavos == 1:
                        texto += ' e um centavo'
                    else:
                        texto += f' e {converter_ate_999(centavos)} centavos'
                return texto
            except:
                return f"{valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

        # Dados do relatório
        processo = dados.get('processo', 'NÃO INFORMADO')
        os_num = dados.get('os', 'NÃO INFORMADO')
        data = dados.get('data', datetime.now().strftime("%d/%m/%Y"))
        imovel = dados.get('imovel', 'NÃO INFORMADO')
        municipio = dados.get('municipio', 'NÃO INFORMADO')
        uf = dados.get('uf', 'GO')
        car = dados.get('car', 'NÃO INFORMADO')
        coordenadas = dados.get('coordenadas', 'NÃO INFORMADO')
        proprietario = dados.get('proprietario', 'NÃO INFORMADO')
        cpf = dados.get('cpf', 'NÃO INFORMADO')
        alertas = dados.get('alertas', 'NÃO INFORMADO')
        intervalo_supressao = dados.get('intervalo_supressao', 'período não informado')

        # Dados das áreas (usando os mesmos da aba Autos e Embargos)
        app_area = dados.get('app_area', '0')
        rl_area = dados.get('rl_area', '0')
        fora_area = dados.get('fora_area', '0')
        dano_area = dados.get('dano_area', '0')

        # Valores das multas
        valor_app = dados.get('valor_app', 'NÃO CALCULADO')
        valor_rl = dados.get('valor_rl', 'NÃO CALCULADO')
        valor_fora = dados.get('valor_fora', 'NÃO CALCULADO')
        valor_dano = dados.get('valor_dano', 'NÃO CALCULADO')

        # Autos e embargos
        auto_app = dados.get('app_auto', '')
        auto_rl = dados.get('rl_auto', '')
        auto_fora = dados.get('fora_auto', '')
        auto_dano = dados.get('dano_auto', '')

        embargo_app = dados.get('app_embargo', '')
        embargo_rl = dados.get('rl_embargo', '')
        embargo_fora = dados.get('fora_embargo', '')

        relatorio = f"""Em cumprimento a Ordem de Serviço SEMAD/GO de nº {os_num} a Gerência de Geoprocessamento e Sensoriamento Remoto (GEGEO) realizou análise geoespacial em torno do(s) alerta(s) de desmatamento de número(s) {alertas} do Mapbiomas, visando monitoramento e tomada de medidas cabíveis sobre o desmatamento no Estado de Goiás.

    No dia {data}, foi realizada fiscalização remota do imóvel rural denominado {imovel} com número de CAR: {car}, coordenadas SIRGAS 2000 ({coordenadas}), município de {municipio}-{uf} e por meio de análise geoespacial foram identificadas supressões, passíveis de autuação, nas seguintes áreas:
        • Área de Preservação Permanente (APP): {app_area} hectares;
        • Área de Reserva Legal (RL): {rl_area} hectares;
        • Área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP): {fora_area} hectares;
        • Dano à vegetação nativa (corte de árvores isoladas): {dano_area} hectares.

    As supressões ocorreram {intervalo_supressao}, e a descrição detalhada de cada polígono encontra-se especificada no Mapa Geral e no Mapa de Análise Temporal anexados a este relatório. 

    No momento da fiscalização, foi constatado que as supressões que deram origem as infrações estão desprovidas de Licenciamento Ambiental e que o Imóvel rural cadastrado no Sistema Nacional de Cadastro Ambiental Rural - SICAR encontra-se em nome de {proprietario}, CPF: {cpf}, o que ensejou a tomada das medidas administrativas cabíveis. Diante de todos esses elementos, originou-se a responsabilização e aplicação dos Autos de Infração e Termos de Embargo instrumentos estes vinculados a este relatório. Tais medidas foram adotadas visando cessar o dano causado ao meio ambiente em que a reparação e/ou compensação pelo dano será estabelecida por meio de medidas técnicas e ambientais no âmbito da licença ambiental do empreendimento ou de termo de compromisso específico.

    {'=' * 60}

    ÁREA DE PRESERVAÇÃO PERMANENTE

    1) - Auto de Infração nº {auto_app} - Estabelecido pelo artigo 43 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 5.000,00 (cinco mil reais) a R$ 50.000,00 (cinquenta mil reais) por hectare ou fração (multa aberta) e para fixação do valor do auto, referente a supressão de {app_area} hectares em Área de Preservação Permanente (APP), foram utilizados os critérios definidos na Orientação Normativa SEMAD Nº 1/2024 considerando o nível de gravidade da infração e identificação da capacidade econômica do infrator.

    GRAVIDADE
    I - Motivo da Infração: pontuação = 10 (intencional, pois realizou a supressão ciente que não tinha autorização para tal);
    II - Consequência para o meio ambiente: pontuação = 30 (moderada, pois supressão de uma Área de Preservação Permanente - APP compromete a função ambiental de preservar os recursos hídricos);
    III - Consequência para a saúde pública ou para socioeconomia da área de abrangência do fato: pontuação = 0 (não foram constatadas evidências de consequência para a saúde pública ou para a socioeconomia da área de abrangências do fato).
    Somatório das etapas = 40 pontos, com classificação Nível B.

    SITUAÇÃO ECONÔMICA
    Não foi possível identificar a capacidade econômica do autuado pela ausência de documentos ou informações, neste caso optou-se pela FAIXA A (Receita Bruta mensal de até 1 salário mínimo, Nível B, mínimo + 0,025% até 0,8% do teto), no entanto, no momento da audiência de autocomposição poderá ser reclassificado conforme a capacidade econômica, mediante comprovação por documentos, conforme § 2º do Art. 8 da ON - 01/2024.

    Diante do exposto chegou-se a seguinte VALORAÇÃO: (R$ 5.000,00 + (0,025% X R$ 50.000,00)) x {app_area} hectares, totalizando o valor de {valor_app}, ficando esta área embargada pelo Termo de Embargo de nº {embargo_app}.

    {'=' * 60}

    RESERVA LEGAL

    2) - Auto de Infração nº {auto_rl} - Estabelecido pelo artigo 51 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 5.000,00 (cinco mil reais) por hectare ou fração (multa fechada), e pela supressão de {rl_area} hectares em área de Reserva Legal (RL) o auto foi valorado da seguinte forma: R$ 5.000,00 x {rl_area} hectares, totalizando o valor de {valor_rl}, ficando esta área embargada pelo Termo de Embargo de nº {embargo_rl}.

    {'=' * 60}

    FORA DE APP E RL

    3) - Auto de Infração nº {auto_fora} - Estabelecido pelo artigo 52 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 1.000,00 (um mil reais) por hectare ou fração (multa fechada), e pela supressão de {fora_area} hectares em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP) o auto foi valorado da seguinte forma: R$ 1.000,00 x {fora_area} hectares, totalizando o valor de {valor_fora}, ficando esta área embargada pelo Termo de Embargo de nº {embargo_fora}.

    {'=' * 60}

    DANO AMBIENTAL

    4) - Auto de Infração nº {auto_dano} – Estabelecido pelo artigo 53 do Decreto nº 6.514/2008 o qual prevê valoração de R$ 300,00 (trezentos reais) por hectare ou fração (multa fechada), e pelo dano (corte de árvores isoladas) de {dano_area} hectares em área fora de Reserva Legal (RL) e fora de Área de Preservação Permanente (APP) o auto foi valorado da seguinte forma: R$ 300,00 x {dano_area} hectares, totalizando o valor de {valor_dano}. Esta área não fica embargada.

    {'-' * 60}
    {municipio}, {data}
    _____________________________________
    Assinatura do Fiscal
    """
        return relatorio

    def obter_valores_multas(self):
        """Obtém os valores das multas de todas as áreas selecionadas"""
        valores = []
        
        if self.check_app and self.check_app.isChecked():
            widget = self.areas_inputs.get("app", {}).get("valor")
            if widget and hasattr(widget, 'text'):
                valor = widget.text().strip()
                if valor:
                    valores.append(f"APP: {valor}")
        
        if self.check_rl and self.check_rl.isChecked():
            widget = self.areas_inputs.get("rl", {}).get("valor")
            if widget and hasattr(widget, 'text'):
                valor = widget.text().strip()
                if valor:
                    valores.append(f"Reserva Legal: {valor}")
        
        if self.check_fora and self.check_fora.isChecked():
            widget = self.areas_inputs.get("fora", {}).get("valor")
            if widget and hasattr(widget, 'text'):
                valor = widget.text().strip()
                if valor:
                    valores.append(f"Área Passível: {valor}")
        
        if self.check_dano and self.check_dano.isChecked():
            widget = self.areas_inputs.get("dano", {}).get("valor")
            if widget and hasattr(widget, 'text'):
                valor = widget.text().strip()
                if valor:
                    valores.append(f"Dano Ambiental: {valor}")
        
        if not valores:
            return "Nenhum valor calculado."
        
        return "\n   ".join(valores)
    
    # ==========================================================
    # MÉTODOS AUXILIARES PARA PRÉ-VISUALIZAÇÃO HTML
    # ==========================================================

    def _formatar_areas_para_html(self, areas_texto):
        """Formata o texto das áreas em HTML bonito - com títulos claros"""
        if "Nenhuma área identificada" in areas_texto:
            return '<div style="color: #6c757d; text-align: center; padding: 20px;">ℹ️ Nenhuma área identificada no imóvel fiscalizado.</div>'

        linhas = areas_texto.split('\n')
        html = []
        area_atual = ""
        dentro_auto = False
        dentro_embargo = False

        for linha in linhas:
            linha = linha.strip()
            if not linha:
                continue

            # Detecta título da área (APP, RL, Fora, Dano)
            if linha.startswith('ÁREA DE PRESERVAÇÃO PERMANENTE') or \
                    linha.startswith('RESERVA LEGAL') or \
                    linha.startswith('ÁREA FORA DE APP E RL') or \
                    linha.startswith('DANO AMBIENTAL') or \
                    linha.startswith('SUPRESSÃO EM ÁREA DE PRESERVAÇÃO PERMANENTE') or \
                    linha.startswith('SUPRESSÃO EM RESERVA LEGAL') or \
                    linha.startswith('SUPRESSÃO EM ÁREA PASSÍVEL') or \
                    linha.startswith('DANO AMBIENTAL - CORTE DE ÁRVORES ISOLADAS'):
                if area_atual:
                    html.append('</div>')
                area_atual = linha
                # Determina a cor baseada no tipo
                cor = "#006b3f"  # verde padrão
                if "APP" in linha or "PRESERVAÇÃO" in linha:
                    cor = "#006b3f"
                elif "RESERVA" in linha:
                    cor = "#006b3f"
                elif "FORA" in linha or "PASSÍVEL" in linha:
                    cor = "#006b3f"
                elif "DANO" in linha:
                    cor = "#006b3f"

                # Icone baseado no tipo
                icone = "📍"
                if "APP" in linha:
                    icone = "🌳"
                elif "RESERVA" in linha:
                    icone = "🌿"
                elif "FORA" in linha or "PASSÍVEL" in linha:
                    icone = "📐"
                elif "DANO" in linha:
                    icone = "⚠️"

                html.append(f'''
                <div class="area-card" style="border-left: 4px solid {cor};">
                    <div class="area-titulo" style="color: {cor};">
                        {icone} {area_atual}
                    </div>
                ''')

            # Detecta AUTO DE INFRAÇÃO
            elif 'AUTO DE INFRAÇÃO' in linha and 'Nº' not in linha:
                dentro_auto = True
                html.append(f'''
                <div style="background-color: #e8f5e9; border-radius: 6px; padding: 10px; margin: 8px 0; border-left: 3px solid #006b3f;">
                    <div style="font-weight: bold; color: #006b3f; font-size: 11pt; margin-bottom: 5px;">📄 AUTO DE INFRAÇÃO</div>
                ''')

            # Detecta EMBARGO
            elif 'EMBARGO' in linha and 'Nº' not in linha:
                dentro_embargo = True
                html.append(f'''
                <div style="background-color: #fff3cd; border-radius: 6px; padding: 10px; margin: 8px 0; border-left: 3px solid #ffc107;">
                    <div style="font-weight: bold; color: #856404; font-size: 11pt; margin-bottom: 5px;">🚫 EMBARGO</div>
                ''')

            # Conteúdo do AUTO
            elif linha.startswith('Por ') and dentro_auto:
                import re
                linha_formatada = re.sub(r'(R\$[\s]*[\d\.,]+(?: a R\$[\s]*[\d\.,]+)?)',
                                         r'<span class="multa-valor">\1</span>', linha)
                html.append(
                    f'<div style="padding: 5px 0 5px 10px; color: #2c3e50; line-height: 1.6; font-size: 10pt;">{linha_formatada}</div>')
                html.append('</div>')
                dentro_auto = False

            # Conteúdo do EMBARGO
            elif linha.startswith('Fica ') and dentro_embargo:
                html.append(
                    f'<div style="padding: 5px 0 5px 10px; color: #2c3e50; line-height: 1.6; font-size: 10pt;">{linha}</div>')
                html.append('</div>')
                dentro_embargo = False

            # Linhas de separação
            elif linha.startswith('---') or linha.startswith('===') or linha.startswith('┌') or linha.startswith(
                    '└') or linha.startswith('│'):
                continue

            # Valor da Multa
            elif linha.startswith('Valor da Multa'):
                import re
                linha_formatada = re.sub(r'(R\$[\s]*[\d\.,]+(?: a R\$[\s]*[\d\.,]+)?)',
                                         r'<span class="multa-valor">\1</span>', linha)
                html.append(
                    f'<div style="padding: 5px 0 5px 10px; font-weight: bold; color: #006b3f;">💰 {linha_formatada}</div>')

            # Outros textos
            elif linha:
                html.append(f'<div style="padding-left: 10px; color: #34495e; font-size: 10pt;">{linha}</div>')

        if area_atual:
            html.append('</div>')

        if not html:
            return '<div style="color: #6c757d; text-align: center; padding: 20px;">ℹ️ Nenhuma área identificada.</div>'

        return '\n'.join(html)

    #def _formatar_artigos_para_html(self, artigos_texto):
        """Formata o texto dos artigos em HTML com títulos claros e estilo unificado"""
        if not artigos_texto or "Nenhum artigo" in artigos_texto:
            return '<div style="color: #6c757d; text-align: center; padding: 20px;">ℹ️ Nenhum artigo selecionado.</div>'

        linhas = artigos_texto.split('\n')
        html = []
        artigo_atual = ""
        dentro_auto = False
        dentro_embargo = False

        for linha in linhas:
            linha = linha.strip()
            if not linha:
                continue

            # Detecta início de um novo artigo
            if linha.startswith('Art.') and ' - ' in linha:
                if artigo_atual:
                    html.append('</div>')
                # Extrai o número do artigo e o nome
                partes = linha.split(' - ')
                num_artigo = partes[0].replace('│', '').strip()
                nome_artigo = partes[1] if len(partes) > 1 else ""
                artigo_atual = linha.replace('│', '').strip()
                html.append(f'''
                <div class="area-card" style="border-left: 4px solid #ff6b00; margin-bottom: 15px; border-radius: 8px; padding: 12px; background-color: #f8f9fa;">
                    <div style="font-weight: bold; font-size: 13pt; color: #ff6b00; margin-bottom: 10px; display: flex; align-items: center; gap: 8px;">
                        📜 {artigo_atual}
                    </div>
                ''')
                dentro_auto = False
                dentro_embargo = False

            # Detecta AUTO DE INFRAÇÃO
            elif 'AUTO DE INFRAÇÃO' in linha:
                dentro_auto = True
                dentro_embargo = False
                html.append(f'''
                <div style="background-color: #e8f5e9; border-radius: 6px; padding: 10px; margin: 8px 0; border-left: 3px solid #006b3f;">
                    <div style="font-weight: bold; color: #006b3f; font-size: 11pt; margin-bottom: 5px;">📄 AUTO DE INFRAÇÃO</div>
                ''')

            # Detecta EMBARGO
            elif 'EMBARGO' in linha:
                dentro_embargo = True
                dentro_auto = False
                html.append(f'''
                <div style="background-color: #fff3cd; border-radius: 6px; padding: 10px; margin: 8px 0; border-left: 3px solid #ffc107;">
                    <div style="font-weight: bold; color: #856404; font-size: 11pt; margin-bottom: 5px;">🚫 EMBARGO</div>
                ''')

            # Conteúdo do AUTO (texto que começa com "Por ")
            elif linha.startswith('Por ') and dentro_auto:
                import re
                linha_formatada = re.sub(r'(R\$[\s]*[\d\.,]+(?: a R\$[\s]*[\d\.,]+)?)',
                                         r'<span style="color: #006b3f; font-weight: bold; font-size: 11pt;">\1</span>',
                                         linha)
                html.append(f'''
                    <div style="padding: 5px 0 5px 10px; color: #2c3e50; line-height: 1.6; font-size: 10pt;">
                        {linha_formatada}
                    </div>
                ''')
                html.append('</div>')
                dentro_auto = False

            # Conteúdo do EMBARGO (texto que começa com "Fica ")
            elif linha.startswith('Fica ') and dentro_embargo:
                html.append(f'''
                    <div style="padding: 5px 0 5px 10px; color: #2c3e50; line-height: 1.6; font-size: 10pt;">
                        {linha}
                    </div>
                ''')
                html.append('</div>')
                dentro_embargo = False

            # Linhas de separação
            elif linha.startswith('---') or linha.startswith('===') or linha.startswith('┌') or linha.startswith(
                    '└') or linha.startswith('│'):
                continue

            # Qualquer outro texto
            elif linha:
                import re
                if 'R$' in linha:
                    linha = re.sub(r'(R\$[\s]*[\d\.,]+(?: a R\$[\s]*[\d\.,]+)?)',
                                   r'<span style="color: #006b3f; font-weight: bold;">\1</span>', linha)
                html.append(f'<div style="padding-left: 10px; color: #34495e; font-size: 10pt;">{linha}</div>')

        if artigo_atual:
            html.append('</div>')

        if not html:
            return '<div style="color: #6c757d; text-align: center; padding: 20px;">ℹ️ Nenhum conteúdo para exibir.</div>'

        return '\n'.join(html)

    #def obter_texto_artigos_selecionados(self):
        """Retorna o texto formatado dos artigos selecionados com títulos claros"""
        textos = []

        artigos_config = {
            "art48": {
                "nome": "Art. 48 - Impedir regeneração de vegetação",
                "template_auto": """Por impedir ou dificultar a regeneração natural de florestas e demais formas de vegetação nativa, na área de {area} hectares, sem autorização do órgão ambiental competente, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.""",
                "template_embargo": """Fica embargada a área de {area} hectares, por impedir ou dificultar a regeneração natural de florestas e demais formas de vegetação nativa, sem autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "art66": {
                "nome": "Art. 66 - Executar atividade sem licença",
                "template_auto": """Por construir, reformar, ampliar, instalar ou fazer funcionar estabelecimentos, atividades, obras ou serviços utilizadores de recursos ambientais, considerados efetiva ou potencialmente poluidores, sem licença ou autorização dos órgãos ambientais competentes, na área de {area} hectares, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.""",
                "template_embargo": """Fica embargada a área de {area} hectares, por executar atividade sem licença ou autorização do órgão ambiental competente, ocorrida no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            },
            "art79": {
                "nome": "Art. 79 - Descumprimento de embargo",
                "template_auto": """Por descumprir embargo instituído por autoridade ambiental competente, reincidindo na atividade embargada, na área de {area} hectares, no interior do imóvel rural denominado {imovel}, município de {municipio} – {uf}, valoração de {valor_multa}.""",
                "template_embargo": """Fica mantido o embargo da área de {area} hectares, por descumprimento de embargo anterior, no imóvel rural denominado {imovel}, município de {municipio} – {uf}."""
            }
        }

        imovel = self.inputs["imovel"].text().strip() if "imovel" in self.inputs and self.inputs[
            "imovel"].text() else "NÃO INFORMADO"
        municipio = self.inputs["municipio"].text().strip() if "municipio" in self.inputs and self.inputs[
            "municipio"].text() else "NÃO INFORMADO"
        uf = self.inputs["uf"].text().strip() if "uf" in self.inputs and self.inputs["uf"].text() else "GO"

        tem_artigo_selecionado = False

        for chave, config in artigos_config.items():
            check = getattr(self, f"check_{chave}", None)
            if check and check.isChecked():
                tem_artigo_selecionado = True

                artigo_data = self.artigos_inputs.get(chave, {})

                area_widget = artigo_data.get("area")
                area_text = area_widget.text().strip() if area_widget and hasattr(area_widget, 'text') else "0"

                valor_widget = artigo_data.get("valor")
                valor_text = valor_widget.text().strip() if valor_widget and hasattr(valor_widget,
                                                                                     'text') else "NÃO CALCULADO"

                # Monta o texto com formatação limpa - SEM títulos repetidos
                texto = f"""
    ┌─────────────────────────────────────────────────────────────────────────────┐
    │ {config['nome']}                                    │
    └─────────────────────────────────────────────────────────────────────────────┘

    📄 AUTO DE INFRAÇÃO
    {config['template_auto'].format(
                    area=area_text,
                    imovel=imovel,
                    municipio=municipio,
                    uf=uf,
                    valor_multa=valor_text
                )}

    🚫 EMBARGO
    {config['template_embargo'].format(
                    area=area_text,
                    imovel=imovel,
                    municipio=municipio,
                    uf=uf
                )}

    """
                textos.append(texto)

        if not tem_artigo_selecionado:
            return ""

        return "\n".join(textos)
    def _get_uc_explicacao(self):
        """Retorna explicação sobre o efeito da UC com o Art. 93"""
        if self.is_uc_marcado():
            return """
            ✓ As multas estão sendo aplicadas em DOBRO conforme Art. 93 do Decreto 6.514/2008.

            📜 Art. 93 do Decreto nº 6.514/2008:
            "As infrações praticadas em Unidade de Conservação ou na sua Zona de Amortecimento 
            terão os valores das multas aplicadas em dobro."

            ⚠️ Importante: Aplica-se também às áreas de preservação permanente (APP) 
            localizadas no interior ou no entorno de Unidades de Conservação.
            """
        else:
            return "ℹ️ Marque esta opção se a infração ocorrer em Unidade de Conservação ou Zona de Amortecimento."
    def _calcular_total_multas(self):
        """Calcula o total das multas de todas as áreas selecionadas"""
        total = 0.0
        for chave in ["app", "rl", "fora", "dano"]:
            if chave in self.areas_inputs:
                valor_widget = self.areas_inputs[chave].get("valor")
                if valor_widget and hasattr(valor_widget, 'text'):
                    valor_text = valor_widget.text().strip()
                    if valor_text:
                        # Remove "R$" e converte
                        import re
                        numeros = re.findall(r'[\d\.,]+', valor_text)
                        if numeros:
                            try:
                                valor_num = float(numeros[0].replace('.', '').replace(',', '.'))
                                total += valor_num
                            except:
                                pass
        if total > 0:
            return f"R$ {total:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
        return "R$ 0,00"
    
    def atualizar_progresso_novos(self, valor):
        self.progress_bar_novos.setValue(valor)
    
    def atualizar_status_novos(self, msg):
        self.label_status_novos.setText(msg)

    def geracao_concluida_novos(self, arquivos):
        self.progress_bar_novos.setVisible(False)
        self.btn_gerar_novos.setEnabled(True)
        self.label_status_novos.setText(f"Concluído! {len(arquivos)} relatório(s) gerado(s).")
        QMessageBox.information(self, "Sucesso", f"{len(arquivos)} relatório(s) gerado(s) com sucesso!")
    
    def erro_geracao_novos(self, erro):
        self.progress_bar_novos.setVisible(False)
        self.btn_gerar_novos.setEnabled(True)
        self.label_status_novos.setText(f"Erro: {erro}")
        QMessageBox.critical(self, "Erro", str(erro))

    def criar_aba_modelos(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        grupo_lista = QGroupBox("Modelos Disponíveis (Combinações de Áreas)")
        lista_layout = QVBoxLayout()
        
        self.lista_modelos = QListWidget()
        self.lista_modelos.setSelectionMode(QListWidget.MultiSelection)
        
        self.modelos_disponiveis = {}
        self.carregar_modelos()
        
        for chave, arquivo in self.modelos_disponiveis.items():
            item = QListWidgetItem(f"{chave.upper()} -> {arquivo}")
            item.setData(Qt.UserRole, chave)
            self.lista_modelos.addItem(item)
        
        lista_layout.addWidget(self.lista_modelos)
        grupo_lista.setLayout(lista_layout)
        layout.addWidget(grupo_lista)
        
        grupo_gerar = QGroupBox("Gerar Relatórios por Combinação de Áreas")
        gerar_layout = QVBoxLayout()
        
        botoes = QHBoxLayout()
        
        self.btn_selecionados = QPushButton("Gerar Selecionados")
        self.btn_auto = QPushButton("Gerar Automático")
        self.btn_limpar = QPushButton("Limpar Tudo")
        
        self.btn_selecionados.clicked.connect(self.gerar_selecionados)
        self.btn_auto.clicked.connect(self.gerar_automatico)
        self.btn_limpar.clicked.connect(self.limpar_tudo)
        
        botoes.addWidget(self.btn_selecionados)
        botoes.addWidget(self.btn_auto)
        botoes.addWidget(self.btn_limpar)
        
        gerar_layout.addLayout(botoes)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        gerar_layout.addWidget(self.progress_bar)
        
        self.label_status = QLabel("Aguardando ação...")
        gerar_layout.addWidget(self.label_status)
        
        grupo_gerar.setLayout(gerar_layout)
        layout.addWidget(grupo_gerar)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget

    def carregar_modelos(self):
        for chave, arquivo in MODELOS.items():
            caminho = os.path.join(PASTA_MODELOS, arquivo)
            if os.path.exists(caminho):
                self.modelos_disponiveis[chave] = arquivo

    def gerar_selecionados(self):
        modelos = []
        for item in self.lista_modelos.selectedItems():
            chave = item.data(Qt.UserRole)
            arquivo = self.modelos_disponiveis.get(chave)
            if arquivo:
                caminho = os.path.join(PASTA_MODELOS, arquivo)
                if os.path.exists(caminho):
                    modelos.append((chave, caminho, chave))
        
        if not modelos:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um modelo!")
            return
        
        self.iniciar_geracao(modelos)

    def gerar_automatico(self):
        tipos = []
        if self.check_app and self.check_app.isChecked(): tipos.append("app")
        if self.check_rl and self.check_rl.isChecked(): tipos.append("rl")
        if self.check_fora and self.check_fora.isChecked(): tipos.append("fora")
        if self.check_dano and self.check_dano.isChecked(): tipos.append("dano")
        
        if not tipos:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos uma área!")
            return
        
        tipos.sort()
        chave = " ".join(tipos)
        arquivo = self.modelos_disponiveis.get(chave)
        
        if not arquivo:
            QMessageBox.warning(self, "Erro", f"Modelo não encontrado para: {chave}")
            return
        
        caminho = os.path.join(PASTA_MODELOS, arquivo)
        if os.path.exists(caminho):
            self.iniciar_geracao([(chave, caminho, chave)])
        else:
            QMessageBox.warning(self, "Erro", f"Arquivo não encontrado: {arquivo}")

    def iniciar_geracao(self, modelos):
        if not self.inputs["processo"].text() or not self.inputs["imovel"].text():
            resposta = QMessageBox.question(
                self, "Aviso",
                "Processo e Imóvel são obrigatórios. Continuar?",
                QMessageBox.Yes | QMessageBox.No
            )
            if resposta == QMessageBox.No:
                return
        
        artigos_valores = {}
        for chave, dados in self.artigos_inputs.items():
            if "valor" in dados:
                artigos_valores[chave] = dados["valor"].text()
            else:
                artigos_valores[chave] = ""
        
        contexto = {
            "processo": self.inputs["processo"].text(),
            "data": datetime.now().strftime("%d/%m/%Y"),
            "imovel": self.inputs["imovel"].text(),
            "municipio": self.inputs["municipio"].text(),
            "car": self.inputs["car"].text(),
            "proprietario": self.inputs["proprietario"].text(),
            "cpf": self.inputs["cpf"].text(),
            "coordenadas": self.inputs["coordenadas"].text(),
            "observacoes": self.obs.toPlainText(),
            "valor_art48": artigos_valores.get("art48", ""),
            "valor_art66": artigos_valores.get("art66", ""),
            "valor_art79": artigos_valores.get("art79", ""),
            "art48_marcado": self.check_art48.isChecked() if self.check_art48 else False,
            "art66_marcado": self.check_art66.isChecked() if self.check_art66 else False,
            "art79_marcado": self.check_art79.isChecked() if self.check_art79 else False,
        }
        
        for chave, dados in self.areas_inputs.items():
            contexto[f"{chave}_area"] = dados["area"].text()
            contexto[f"{chave}_auto"] = dados["auto"].text()
            contexto[f"{chave}_embargo"] = dados["embargo"].text()
            contexto[f"{chave}_valor"] = dados["valor"].text()
        
        for chave, dados in self.artigos_inputs.items():
            contexto[f"{chave}_auto"] = dados["auto"].text()
            contexto[f"{chave}_embargo"] = dados["embargo"].text()
        
        pasta = QFileDialog.getExistingDirectory(self, "Escolha a pasta para salvar")
        if not pasta:
            return
        
        self.thread = GeradorRelatoriosThread(modelos, contexto, pasta)
        self.thread.progresso.connect(self.atualizar_progresso)
        self.thread.status.connect(self.atualizar_status)
        self.thread.concluido.connect(self.geracao_concluida)
        self.thread.erro.connect(self.erro_geracao)
        
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.btn_selecionados.setEnabled(False)
        self.btn_auto.setEnabled(False)
        
        self.thread.start()

    def atualizar_progresso(self, valor):
        self.progress_bar.setValue(valor)

    def atualizar_status(self, msg):
        self.label_status.setText(msg)

    def geracao_concluida(self, arquivos):
        self.progress_bar.setVisible(False)
        self.btn_selecionados.setEnabled(True)
        self.btn_auto.setEnabled(True)
        self.label_status.setText(f"Concluído! {len(arquivos)} relatório(s) gerado(s).")
        QMessageBox.information(self, "Sucesso", f"{len(arquivos)} relatório(s) gerado(s)!")

    def erro_geracao(self, erro):
        self.progress_bar.setVisible(False)
        self.btn_selecionados.setEnabled(True)
        self.btn_auto.setEnabled(True)
        self.label_status.setText(f"Erro: {erro}")
        QMessageBox.critical(self, "Erro", str(erro))

    def limpar_tudo(self):
        for campo in self.inputs.values():
            campo.clear()
        
        for area in self.areas_inputs.values():
            area["area"].clear()
            area["auto"].clear()
            area["embargo"].clear()
            area["valor"].clear()
        
        for artigo in self.artigos_inputs.values():
            artigo["area"].clear()
            artigo["auto"].clear()
            artigo["embargo"].clear()
            artigo["valor"].clear()
            if "motivo" in artigo:
                artigo["motivo"].setCurrentIndex(0)
                artigo["consequencia"].setCurrentIndex(0)
                artigo["saude"].setCurrentIndex(0)
                artigo["pontuacao"].clear()
                artigo["faixa_receita"].setCurrentIndex(0)
                artigo["percentual"].setText("Percentual aplicado: -")
        
        self.obs.setText("")
        
        if self.check_app:
            self.check_app.setChecked(False)
        if self.check_rl:
            self.check_rl.setChecked(False)
        if self.check_fora:
            self.check_fora.setChecked(False)
        if self.check_dano:
            self.check_dano.setChecked(False)
        
        if self.check_art48:
            self.check_art48.setChecked(False)
        if self.check_art66:
            self.check_art66.setChecked(False)
        if self.check_art79:
            self.check_art79.setChecked(False)
        
        self.label_status.setText("Aguardando ação...")


# ==========================================================
# EXECUÇÃO
# ==========================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = SistemaRelatorios()
    janela.show()
    sys.exit(app.exec())