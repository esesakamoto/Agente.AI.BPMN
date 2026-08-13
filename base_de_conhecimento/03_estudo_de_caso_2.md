# Estudo de Caso 2 – Aplicabilidade da Análise de Séries Temporais para Melhoria nos Indicadores de Desempenho da SES-DF

APLICABILIDADE DA ANÁLISE DE SÉRIES TEMPORAIS PARA MELHORIA NOS INDICADORES DE DESEMPENHO DA SES-DF MARILZA OLIVEIRA DE ALMEIDA1 CHRISTIANE BRAGA MARTINS DE BRITO2 GRACIELA PAULI GIL CARDOSO3 CYNTHIA RODOVALHO ROSA4 GERALDO MAGELA SARAIVA GAMA5

1. Apresentação

Este artigo tem por objetivo apresentar a arquitetura e a criação de uma ferramenta na planilha eletrônica Microsoft Excel, aplicada ao processo de análise de séries temporais ou históricas de um conjunto de indicadores de saúde, 1 Gestora de Políticas Públicas e Gestão Governamental-Administradora, mestre em Avaliação em Saúde, especialista em Avaliação em Saúde, Gestão do Trabalho e Educação em Saúde, lotada na SUPLANS/SES-DF. 2 Subsecretária de Planejamento em Saúde (SUPLANS) da /SES-DF, enfermeira, especialista em Controle de Infecção, especialista em Gestão de Negócios e mestranda em Avaliação em Saúde. 3 Diretora da Diretoria de Planejamento (DIPLAN), da SUPLANS/SES-DF, assistente social, mestre em Gestão do Conhecimento e Tecnologia da Informação e especialista em Avaliação em Saúde. 4 Técnica em Saúde, Matemática, especialista em Estatística Aplicada, lotada na Gerência de Monitoramento e Avaliação da DIPLAN/SUPLANS/SES-DF. 5 Especialista em Saúde-Administrador, especialista em Gestão Pública, lotado na Diretoria de Gestão de Informações Estratégicas da SUPLANS/SES-DF. 17

embasada em métodos e técnicas que apoiem e auxiliem gestores e profissionais a definirem metas.

A ferramenta permite, por meio do registro histórico dos indicadores, mensurar a tendência atual, fazer estimativas, descrever o comportamento da série e, por fim, avaliar quais os fatores que influenciaram o comportamento da série, além de revisar as metas negociadas. Para que isso ocorresse, foram utilizadas funções matemáticas, técnicas estatísticas, como o Controle Estatístico de Processo (CEP), passando a integrar o ciclo de planejamento geral da Secretaria de Estado de Saúde do Distrito Federal (SES-DF) por meio dos instrumentos de planejamento para o quadriênio (2020-2023): o Plano Plurianual (PPA), o Plano Distrital de Saúde (PDS) e outras pactuações governamentais.

Ressalta-se, ainda, que essa foi uma construção participativa das análises de séries temporais que propiciaram o compartilhamento de experiências, conhecimentos e habilidades, oriundos dos técnicos que buscaram o melhor de si em favor do alcance da melhoria dos processos organizacionais.

2. Contextualização

2.1. A organização

A gestão do Sistema Único de Saúde do Distrito Federal (SUS-DF) compete a SES-DF, responsável pela organização e elaboração de planos, programas, projetos e políticas públicas voltadas para a promoção, prevenção e assistência à saúde (DISTRITO FEDERAL, 2019). Tem como missão garantir ao cidadão acesso universal à saúde mediante atenção integral e humanizada.

Em 2016, por meio do Decreto nº 37.515, foi instituído o Programa de Gestão Regional da Saúde (PRS) para as Regiões de Saúde e Unidades de Referência Distrital (URDs). O DF compõe sete Regiões de Saúde: Central, Centro-Sul, Oeste, Sul, Sudoeste, Norte e Leste e cinco URDs compostas pelo Hospital de Base do Distrito Federal (HBDF), Hospital Materno Infantil (HMIB), Hospital São Vicente de Paulo (HSVP), Hospital de Apoio de Brasília (HAB), Hospital da Criança de Brasília José Alencar (HCB) ( DISTRITO FEDERAL, 2016; 2018).

A SES-DF possui uma estrutura complexa, ampla e abrangente, estando organizada em um sistema regionalizado e hierarquizado por meio de uma rede de serviço composta pela atenção primária, atenção secundária, atenção especializada ambulatorial e hospitalar, atenção psicossocial, serviços de

urgência e emergência, serviços das vigilâncias sanitária, epidemiológica e ambiental e assistência farmacêutica.

Para monitorar e avaliar suas ações, serviços e produtos ofertados à população, a SES-DF, em 2016, desenvolveu uma ferramenta gerencial dos instrumentos de planejamento, inicialmente em plataforma Excel. Em 2018, passou a tecnologia livre e em plataforma web, chamada Sistema de Planejamento da SES-DF (SESPLAN). Esta ferramenta permite o registro de dados das diversas áreas da secretaria, compartilhando as informações do ciclo do planejamento, ampliando a compreensão dos macroprocessos, desde os finalísticos até os de sustentação e a análise dos resultados (CARDOSO et al., 2018; SELLERA et al., 2019).

A partir daí, com o ciclo de planejamento estabelecido, o sistema SESPLAN e a sistemática de M&A implantados, idealizou-se o passo seguinte para a melhoria do processo de definição de metas para indicadores, com o objetivo de analisar as séries histórias com a ajuda de software de planilha eletrônica Excel, utilizando as funções matemáticas, os métodos estatísticos e o Controle de Estatístico de Qualidade (CEP), este mundialmente conhecido após os anos 1960, com a aplicação do ciclo PDCA.

2.2. Diagnóstico do problema

As diretrizes do processo de planejamento no âmbito do SUS têm como base os seguintes pressupostos, entre outros: o planejamento como responsabilidade individual de cada um dos três entes federados (municípios, estados e DF); monitoramento, avaliação e integração da gestão do SUS e compatibilização entre os instrumentos de planejamento da saúde (Plano de Saúde e respectivas Programações Anuais, Relatórios de Gestão); e os instrumentos de planejamento e orçamento de governo, quais sejam o Plano Plurianual (PPA), a Lei de Diretrizes Orçamentárias (LDO) e a Lei Orçamentária Anual (LOA), em cada esfera de gestão (BRASIL, 2013; 2017).

Neste contexto, para a elaboração dos instrumentos de planejamento da SES-DF para o quadriênio 2020-2023, identificou-se a necessidade de alinhar as metas e os indicadores, uma vez que as distorções existentes e os desencontros de metas ora superestimadas, ora subestimadas pelas áreas técnicas, ignorando em grande parte as séries históricas de desempenho, dificultavam o alcance dos objetivos estratégicos pretendidos. Além disso, percebeu-se

também que havia uma desinformação e/ou confusão sobre objetivos, metas e indicadores por parte de alguns representantes das áreas finalísticas.

Picchiai (2012) afirma que as metas e os indicadores atuam como elementos da materialização da estratégia. As metas são quantificação dos objetivos; os indicadores são a parametrização e numeração das atividades desenvolvidas no processo de implantação da estratégia e do planejamento.

Logo, foi necessário fazer um nivelamento desses conceitos, o planejamento de metas para os indicadores, por meio de análise de séries temporais, a fim de servirem de alerta confiável para tomada de decisão e de mudança de estratégias pelos gestores.

Para Everitt (1995) e Morretin e Tolloi (2006) uma série histórica é uma sequência de dados obtidos em intervalos regulares de tempo durante um período específico.

Miranda (2014) define série temporal como um conjunto de observações de uma determinada variável feitas em períodos sucessivos de tempo, ao longo de um determinado intervalo.

O CEP ou Carta de Controle é definido por Ribeiro e Caten (2012) como uma técnica estatística aplicada à produção que permite a redução sistemática da variabilidade nas características da qualidade de interesse, além de contribuir para a melhoria da qualidade intrínseca, da produtividade, da confiabilidade e do custo do que está sendo produzido.

Segundo Schultz (2019), o CEP é um método que controla os resultados de um fluxo de produção, por meio de indicadores que garantam a qualidade do produto, tendo como principal objetivo padronizar e estabilizar o processo, de forma a sustentar as melhorias.

Dessa forma, a motivação para criação de uma ferramenta que possibilitasse a pactuação de metas para os indicadores, utilizando a análise da série histórica ou série temporal com a combinação de funções matemáticas, métodos e técnicas estatísticas foi idealizada e implementada na SES-DF.

3. Construção da solução

A ferramenta foi desenvolvida no software de planilhas eletrônicas Microsoft Excel, em março de 2019, nomeada pela equipe idealizadora de Controle de

Tendência de Resultados (CTR ), para auxiliar as áreas técnicas na pactuação de metas para os indicadores estratégicos com a aplicação da combinação do Controle Estatístico do Processo (CEP), da Análise de Série Temporal (ASI) e da Linha de Tendência (LT).

Para utilização da ferramenta CTR são necessários os dados das séries históricas dos indicadores, metas pactuadas e resultados alcançados e, assim, geram-se as projeções para o período definido, baseado no comportamento dos dados de anos anteriores.

Em seguida, analisa-se o comportamento dos dados ao longo do tempo, por meio estatístico, para então prever a série temporal, ou seja, a determinação dos prováveis valores que assumirão as variáveis futuras; o que servirá de base técnica para decisões atuais e futuras, conforme pode ser demonstrado na Figura 2.1.

Figura 2.1: Representação da Análise de série Temporal.

Fonte: CTR, março, 2019.

A Figura 2.2 demonstra a aplicação do CEP.

Figura 2.2: Representação gráfica do CEP ou Carta de Controle.

Fonte: CTR, março, 2019.

A Figura 3.2 é formada por três linhas paralelas: a central que representa o resultado médio encontrado por meio dos valores coletados e um par de limites de controle, um destes se localiza abaixo e outro acima da linha central, determinando assim o limite inferior de controle (LIC) e o limite superior de controle (LSC). Um processo estará sob controle quando todos os pontos estiverem posicionados dentro do limite inferior e superior, entretanto, quanto mais os pontos estiverem posicionados próximos à linha central, mais confiável é o processo (PYLRO, 2008).

Na SES/DF adotou-se a projeção utilizando a Linha de Tendência do tipo linear, em razão do processo de amadurecimento institucional e das limitações da ferramenta para adotar todos os modelos disponíveis como: Exponencial, Logarítmica, Polinomial, Potência ou Média Móvel.

Em todas as análises produzidas são inseridas a equação no gráfico e o valor de R-Quadrado, conforme a Figura 2.3.

Figura 2.3: Demonstrativo da equação e R-quadrado.

Fonte: CTR, junho, 2019.

O R-Quadrado funciona como parâmetro de ajustes da negociação das metas dos indicadores com a área técnica. O R-Quadrado é uma medida estatística de quão próximos os dados estão da linha de regressão ajustada.

É também conhecido como o coeficiente de determinação ou o coeficiente de determinação múltipla para a regressão múltipla ou R-quadrado ajustado.

Varia entre 0 e 1 (quanto mais próximo de 1, melhor a linha de tendência) (NASCIMENTO e ARAÚJO, 2009).

O processo de envolvimento dos interessados ocorreu na oficina do planejamento estratégico com área técnica e gestores da SES-DF, oportunidade em que se demonstrou na prática e pôde ser validada a utilização da ferramenta, conforme breve relato.

Nas primeiras oficinas identificaram-se os cenários, utilizando da estratégia com foco no usuário e na Rede de Atenção à Saúde (RAS), cujo produto foi a identificação das necessidades do usuário e os entraves dos serviços, assim como se trabalhou a Cadeia de Valor, Visão de Futuro e o Balanced Scorecard (BSC), identificando diretrizes e objetivos. Para finalizar, após alinhar os objetivos, foi apresentada a seleção de indicadores estratégicos para PDS e PPA, bem como a proposta de metas para os indicadores selecionados.

Nesse momento, mostrou-se por meio da ferramenta que algumas metas não eram compatíveis com os seus indicadores. Ao final, os níveis de desempenho pretendidos para o quadriênio (2020-2023) eram chancelados e as áreas percebiam se estavam perto ou distante de atingirem seus objetivos, com os aportes existentes e previstos.

O processo da análise da série temporal realizado na ferramenta CTR levou em consideração as seguintes etapas: • seleção dos indicadores com suas respectivas metas e resultados; • interpretação do comportamento dos dados ao longo do tempo; • cálculo dos limites de controle; • interpretação da estabilidade do processo; • interpretação da capacidade do processo; • definição das metas compatíveis para o indicador; • pactuação das metas dos indicadores com as áreas.

A representação do resultado final da análise utilizando a ferramenta CTR-SES-DF é demonstrada na Figura 2.4:

Figura 2.4: Demonstração da ferramenta CTR/SES-DF, após aplicação de um indicador.

Legenda: laranja: metas; azul-claro: resultados; verde: média aritmética simples; azul-escuro:

LSC (Limite Superior de Controle); vermelho: LIC (Limite Inferior de Controle) e rosa: parâmetro.

Fonte: CTR, junho, 2019.

Na etapa de pactuação das metas dos indicadores com as áreas finalísticas, a ferramenta CTR-SES-DF gera o relatório de pactuação, contendo número da pactuação, nome do indicador, área, valores pactuados com as respectivas assinaturas, conforme mostra a Figura 2.5.

Figura 2.5: Relatório de Pactuação de metas gerados pelo CTR-SES-DF.

Fonte: CTR, junho, 2019.

4. Resultados

A ferramenta CTR/SES-DF surgiu como uma solução para as dificuldades de se estabelecer metas confiáveis para indicadores estratégicos, a partir de funções matemáticas, séries temporais, controle estatístico, como já mencionados anteriormente.

Os resultados apontam melhores condições de avaliação das tendências dos indicadores; a comparação dos resultados entre períodos; provimento de informações de advertência; antecipação de futuras condições e tendências demonstrando conformidades e não conformidades das metas para os indicadores, o que reforça o seu caráter inovador, propiciando discussões para o alcance de metas factíveis.

Tratando-se de uma planilha Excel, esta ferramenta possui limitações.

No entanto, para a melhoria contínua foram definidas as regras do negócio a fim de que o CTR se integre como módulo a ser desenvolvido e customizado no sistema SESPLAN, o que se constituirá um grande avanço tecnológico para a SES-DF.

Por último, estudos subsequentes devem ser conduzidos para a apropriação de possibilidades das combinações de métodos e técnicas estatísticas que ainda podem ser desenvolvidos no CTR.

Referências

BRASIL. MINISTÉRIO DA SAÚDE. Portaria no 2.135, de 25 de setembro de 2013.

Estabelece diretrizes para o processo de planejamento no âmbito do Sistema Único de Saúde (SUS). Disponível em: <http://bvsms.saude.gov.br/bvs/saudelegis/ gm/2013/prt2135_25_09_2013.html#:~:text=PORTARIA%20N%C2%BA%20 2.135%2C%20DE%2025,%C3%9Anico%20de%20Sa%C3%BAde%20(SUS).&text=I%20 %2D%20planejamento%20como%20responsabilidade%20individual,forma%20 cont%C3%ADnua%2C%20articulada%20e%20integrada>. Acesso em: 15 jun. 2020.

BRASIL. MINISTÉRIO DA SAÚDE. Portaria de Consolidação no 1, de 28 de setembro de 2017. Consolidação das normas sobre os direitos e deveres dos usuários da saúde, a organização e o funcionamento do Sistema Único de Saúde. Disponível em: <http:// portalsinan.saude.gov.br/images/documentos/Legislacoes/Portaria_Consolidacao_ 1_28_SETEMBRO_2017.pdf>. Acesso em: 22 dez. 2017.

DISTRITO FEDERAL. Decreto no 37.057, de 14 de janeiro de 2016. Dispõe sobre a estrutura administrativa da Secretaria de Estado de Saúde do Distrito Federal, que

especifica e dá outras providências. Disponível em: http:// www.sinj.df.gov.br/sinj/ Norma/ef9ffafc95b0448db85e0010846badf8/exec_dec_37057_2016_rep.html. Acesso em: 04 jul. 2020.

DISTRITO FEDERAL. Decreto no 38.982, de 10 de abril de 2018. Altera a estrutura administrativa da Secretaria de Estado de Saúde do Distrito Federal e dá outras providências.

Disponível em: http://www.saude.df.gov.br/wp-conteudo/uploads/2018/04/Decreton.%C2%BA-38.982-2018-Altera-a-estrutura-adminsitrativa-da-SES-DF.pdf. Acesso em: 10 jun. 2020.

DISTRITO FEDERAL. Decreto no 37.515, de 26 de julho de 2016. Institui o Programa de Gestão Regional da Saúde - PRS para as Regiões de Saúde e Unidades de Referência Distrital. Disponível em: <http://www.sinj.df.gov.br/sinj/Norma/05990c188e6a4778 860953ca699e356e/exec_dec_37515_2016.html#:~:text=DECRETO%20N%C2%BA%20 37.515%2C%20DE%2026,e%20Unidades%20de%20Refer%C3%AAncia%20Distrital>.

Acesso em: 18 jun. 2020.

DISTRITO FEDERAL. Decreto no 39.610, de 1o de janeiro de 2019. Dispõe sobre a organização da estrutura da Administração Pública do Distrito Federal. Disponível em: <http://www.buriti.df.gov.br/ftp/diariooficial/2019/01_Janeiro/DODF%20001%2001-012019 %20EDICAO%20ESPECIAL/DODF%20001%2001-01-2019%20EDICAO%20ESPECIAL. pdf>. Acesso em: 12 jun. 2020.

DISTRITO FEDERAL. Lei no 6.270, de 30 de janeiro de 2019. Altera a nomenclatura do Instituto Hospital de Base do Distrito Federal – IHBDF, instituído pela Lei nº 5.899, de 3 de julho de 2017, para Instituto de Gestão Estratégica de Saúde do Distrito Federal - IGESDF e dá outras providências. Disponível em: http://www.tc.df.gov.br/sinj/Norma/309ebad 7cfa346c3991cde576e6a57cf/Lei_6270_30_01_2019.html. Acesso em: 07 jun.2020.

CARDOSO, G. J. D. P. G. et al. Construção de um modelo para a gestão do ciclo do planejamento integrado da Secretaria de Estado de Saúde do Distrito Federal: Relato de Experiência. Ciências Saúde. 2018; 29 Supl 1:18-24 <http://www.escs.edu.br/revistaccs/index.php/comunicacaoemcienciasdasaude/ article/view/182/246>. Acesso em: 07 de jul. 2020.

EVERITT, B.S. The Cambridge dictionary of statistics in the medical sciences. Cambridge:

Cambridge University Press; 1995. Disponível em: <http://www.stewartschultz.com/statistics /books/Cambridge%20Dictionary%20Statistics%204th.pdf>. Acesso em: 20 ago 2019.

MIRANDA, I. P. De. Comparação de diferentes Métodos de Previsão em Séries Temporais com valores discrepantes. Monografia (Bacharelado em Estatística) – Universidade Federal de Juiz de Fora, UFJF/MG, 2014.

MORETTIN, P.A; TOLLOI, C.M.C. Previsão de séries temporais. 2a. ed. revisada e ampliada. São Paulo: Blucher; 2006. Disponível em: <https://issuu.com/editorablucher/ docs/issuu_analise_temporais_isbn9788521203896>. Acesso em: 15 jun.2019.

NASCIMENTO, G.; ARAUJO, P. F. Estudo acerca do coeficiente de determinação nos modelos lineares e algumas generalizações, 2009. Disponível em: <https://docs.ufpr. br/~lucambio/CE229/TCC_Patricia_e_Gisele.pdf>. Acesso em: 11 mar. 2019.

PICCHIAI, D. As metas e os indicadores no processo de planejamento: o caso de uma universidade pública. XXXII Encontro Nacional de Engenharia de Produção Desenvolvimento Sustentável e Responsabilidade Social: As Contribuições da Engenharia de Produção Bento Gonçalves, RS, Brasil, 15 a 18 de outubro de 2012. Disponível em: <http://www.abepro.org.br/biblioteca/enegep2012_TI_ST_163_949_21020.pdf>.

Acesso em: 10 jun. 2019.

PYLRO, A. S. Modelo linear dinâmico de Harrison & Stevens aplicado ao controle estatístico de processos autocorrelacionados. Tese (Mestrado em Engenharia de Produção) Pontifícia Universidade Católica do Rio de Janeiro, Rio de Janeiro, 2008. Disponível em:<http://livros01.livrosgratis.com.br/cp077091.pdf>. Acesso em:

RIBEIRO, J. L.; CATEN, C. S. T. Controle estatístico do processo. Série monográfica.

Programa de Pós-graduação em Engenharia de Produção. Porto Alegre: Universidade Federal do Rio Grande do Sul, 2012. Disponível em: <http://www.producao.ufrgs.br/ arquivos/disciplinas/388_apostilacep_2012.pdf>. Acesso em: 18 abr. 2019.

SCHULTZ, F. Controle estatístico de processo: o que é o CEP e para que serve? Disponível em: <https://bomcontrole.com.br/controle-estatistico-processo-cep/>. Acesso em: 10 dez. 2019.

SELLERA, P. E. G. et. al. A implantação do sistema de monitoramento e avaliação da Secretaria Estadual de Saúde do Distrito Federal (SES/DF). Ciênc. saúde coletiva, vol.24, no.6. Rio de Janeiro, Jun. ,2019. Epub Jun. 27, 2019. Disponível em: <https://www.scielo.br/pdf/csc/v24n6/1413-8123-csc-24-06-2085.pdf>. Acesso em: 07 jul. 2020.
