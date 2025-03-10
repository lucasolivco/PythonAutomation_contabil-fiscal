import json
import sqlite3
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.chrome.options import Options
from tkinter import ttk, messagebox
import tkinter as tk
import re
import threading
import os

#INSTAlE pip install selenium beautifulsoup4 webdriver_manager

# Definir o diretório do banco de dados
db_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'CNPJApp')
os.makedirs(db_dir, exist_ok=True)
db_path = os.path.join(db_dir, 'cnae_database.db')

# Criar ou conectar ao banco de dados
try:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
except sqlite3.Error as e:
    print(f"Erro ao conectar ao banco de dados: {e}")
    exit(1)

try:
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS cnae_fator_r (
        cnae TEXT PRIMARY KEY,
        descricao TEXT,
        fator_r BOOLEAN
    )
    ''')
    conn.commit()
except sqlite3.Error as e:
    print(f"Erro ao criar tabela: {e}")
    conn.close()
    exit(1)

# Função para popular o banco de dados
def populate_database():
    cnae_data = [
        ('3250706', 'Serviços de prótese dentária', True),
        ('3250709', 'Serviço de laboratório óptico', True),
        ('4399101', 'Administração de obras', False),
        ('4512901', 'Representantes comerciais e agentes do comércio de veículos automotores', True),
        ('4530702', 'Comércio por atacado de pneumáticos e câmaras-de-ar', False),
        ('4530703', 'Comércio a varejo de peças e acessórios novos para veículos automotores', False),
        ('4530704', 'Comércio a varejo de peças e acessórios usados para veículos automotores', False),
        ('4530705', 'Comércio a varejo de pneumáticos e câmaras-de-ar', False),
        ('4530706', 'Representantes comerciais e agentes do comércio de peças e acessórios novos e usados para veículos automotores', True),
        ('4541202', 'Comércio por atacado de peças e acessórios para motocicletas e motonetas', False),
        ('4541206', 'Comércio a varejo de peças e acessórios novos para motocicletas e motonetas', False),
        ('4541207', 'Comércio a varejo de peças e acessórios usados para motocicletas e motonetas', False),
        ('4542101', 'Representantes comerciais e agentes do comércio de motocicletas e motonetas, peças e acessórios', True),
        ('4611700', 'Representantes comerciais e agentes do comércio de matérias-primas agrícolas e animais vivos', True),
        ('4612500', 'Representantes comerciais e agentes do comércio de combustíveis, minerais, produtos siderúrgicos e químicos', True),
        ('4613300', 'Representantes comerciais e agentes do comércio de madeira, material de construção e ferragens', True),
        ('4614100', 'Representantes comerciais e agentes do comércio de máquinas, equipamentos, embarcações e aeronaves', True),
        ('4615000', 'Representantes comerciais e agentes do comércio de eletrodomésticos, móveis e artigos de uso doméstico', True),
        ('4616800', 'Representantes comerciais e agentes do comércio de têxteis, vestuário, calçados e artigos de viagem', True),
        ('4617600', 'Representantes comerciais e agentes do comércio de produtos alimentícios, bebidas e fumo', True),
        ('4618401', 'Representantes comerciais e agentes do comércio de medicamentos, cosméticos e produtos de perfumaria', True),
        ('4618402', 'Representantes comerciais e agentes do comércio de instrumentos e materiais odonto-médico-hospitalares', True),
        ('4618403', 'Representantes comerciais e agentes do comércio de jornais, revistas e outras publicações', True),
        ('4618499', 'Outros representantes comerciais e agentes do comércio especializado em produtos não especificados anteriormente', True),
        ('4619200', 'Representantes comerciais e agentes do comércio de mercadorias em geral não especializado', True),
        ('4711301', 'Comércio varejista de mercadorias em geral, com predominância de produtos alimentícios - hipermercados', False),
        ('4711302', 'Comércio varejista de mercadorias em geral, com predominância de produtos alimentícios - supermercados', False),
        ('4712100', 'Comércio varejista de mercadorias em geral, com predominância de produtos alimentícios - minimercados, mercearias e armazéns', False),
        ('4713002', 'Lojas de variedades, exceto lojas de departamentos ou magazines', False),
        ('4713004', 'Lojas de departamentos ou magazines, exceto lojas francas (Duty free)', False),
        ('4721102', 'Padaria e confeitaria com predominância de revenda', False),
        ('4721103', 'Comércio varejista de laticínios e frios', False),
        ('4721104', 'Comércio varejista de doces, balas, bombons e semelhantes', False),
        ('4723700', 'Comércio varejista de bebidas', False),
        ('4724500', 'Comércio varejista de hortifrutigranjeiros', False),
        ('4729601', 'Tabacaria', False),
        ('4729699', 'Comércio varejista de produtos alimentícios em geral ou especializado em produtos alimentícios não especificados anteriormente', False),
        ('4743100', 'Comércio varejista de vidros', False),
        ('4744002', 'Comércio varejista de madeira e artefatos', False),
        ('4751201', 'Comércio varejista especializado de equipamentos e suprimentos de informática', False),
        ('4751202', 'Recarga de cartuchos para equipamentos de informática', False),
        ('4752100', 'Comércio varejista especializado de equipamentos de telefonia e comunicação', False),
        ('4753900', 'Comércio varejista especializado de eletrodomésticos e equipamentos de áudio e vídeo', False),
        ('4754701', 'Comércio varejista de móveis', False),
        ('4754702', 'Comércio varejista de artigos de colchoaria', False),
        ('4754703', 'Comércio varejista de artigos de iluminação', False),
        ('4755501', 'Comércio varejista de tecidos', False),
        ('4755502', 'Comércio varejista de artigos de armarinho', False),
        ('4755503', 'Comércio varejista de artigos de cama, mesa e banho', False),
        ('4756300', 'Comércio varejista especializado de instrumentos musicais e acessórios', False),
        ('4757100', 'Comércio varejista especializado de peças e acessórios para aparelhos eletroeletrônicos para uso doméstico, exceto informática e comunicação', False),
        ('4759801', 'Comércio varejista de artigos de tapeçaria, cortinas e persianas', False),
        ('4759899', 'Comércio varejista de outros artigos de uso doméstico não especificados anteriormente', False),
        ('4761001', 'Comércio varejista de livros', False),
        ('4761002', 'Comércio varejista de jornais e revistas', False),
        ('4761003', 'Comércio varejista de artigos de papelaria', False),
        ('4762800', 'Comércio varejista de discos, CDs, DVDs e fitas', False),
        ('4763601', 'Comércio varejista de brinquedos e artigos recreativos', False),
        ('4763602', 'Comércio varejista de artigos esportivos', False),
        ('4763603', 'Comércio varejista de bicicletas e triciclos; peças e acessórios', False),
        ('4763604', 'Comércio varejista de artigos de caça, pesca e camping', False),
        ('4772500', 'Comércio varejista de cosméticos, produtos de perfumaria e de higiene pessoal', False),
        ('4774100', 'Comércio varejista de artigos de óptica', False),
        ('4781400', 'Comércio varejista de artigos do vestuário e acessórios', False),
        ('4782201', 'Comércio varejista de calçados', False),
        ('4782202', 'Comércio varejista de artigos de viagem', False),
        ('4783102', 'Comércio varejista de artigos de relojoaria', False),
        ('4785701', 'Comércio varejista de antiguidades', False),
        ('4785799', 'Comércio varejista de outros artigos usados', False),
        ('4789001', 'Comércio varejista de suvenires, bijuterias e artesanatos', False),
        ('4789002', 'Comércio varejista de plantas e flores naturais', False),
        ('4789003', 'Comércio varejista de objetos de arte', False),
        ('4789007', 'Comércio varejista de equipamentos para escritório', False),
        ('4789008', 'Comércio varejista de artigos fotográficos e para filmagem', False),
        ('4789099', 'Comércio varejista de outros produtos não especificados anteriormente', False),
        ('5611201', 'Restaurantes e similares', False),
        ('5611203', 'Lanchonetes, casas de chá, de sucos e similares', False),
        ('5611204', 'Bares e outros estabelecimentos especializados em servir bebidas, sem entretenimento', False),
        ('5620101', 'Fornecimento de alimentos preparados preponderantemente para empresas', False),
        ('5620103', 'Cantinas - serviços de alimentação privativos', False),
        ('5620104', 'Fornecimento de alimentos preparados preponderantemente para consumo domiciliar', False),
        ('5911101', 'Estúdios cinematográficos', False),
        ('5911102', 'Produção de filmes para publicidade', False),
        ('5911199', 'Atividades de produção cinematográfica, de vídeos e de programas de televisão não especificadas anteriormente', False),
        ('5912001', 'Serviços de dublagem', False),
        ('5912002', 'Serviços de mixagem sonora em produção audiovisual', False),
        ('5912099', 'Atividades de pós-produção cinematográfica, de vídeos e de programas de televisão não especificadas anteriormente', False),
        ('5913800', 'Distribuição cinematográfica, de vídeo e de programas de televisão', False),
        ('5914600', 'Atividades de exibição cinematográfica', False),
        ('5920100', 'Atividades de gravação de som e de edição de música', False),
        ('6201501', 'Desenvolvimento de programas de computador sob encomenda', True),
        ('6201502', 'Web design', True),
        ('6202300', 'Desenvolvimento e licenciamento de programas de computador customizáveis', True),
        ('6203100', 'Desenvolvimento e licenciamento de programas de computador não-customizáveis', True),
        ('6204000', 'Consultoria em tecnologia da informação', True),
        ('6209100', 'Suporte técnico, manutenção e outros serviços em tecnologia da informação', True),
        ('6311900', 'Tratamento de dados, provedores de serviços de aplicação e serviços de hospedagem na internet', True),
        ('6319400', 'Portais, provedores de conteúdo e outros serviços de informação na internet', True),
        ('6391700', 'Agências de notícias', False),
        ('6399200', 'Outras atividades de prestação de serviços de informação não especificadas anteriormente', False),
        ('6619302', 'Correspondentes de instituições financeiras', False),
        ('6621501', 'Peritos e avaliadores de seguros', True),
        ('6621502', 'Auditoria e consultoria atuarial', True),
        ('6622300', 'Corretores e agentes de seguros, de planos de previdência complementar e de saúde', False),
        ('6629100', 'Atividades auxiliares dos seguros, da previdência complementar e dos planos de saúde não especificadas anteriormente', False),
        ('6821801', 'Corretagem na compra e venda e avaliação de imóveis', False),
        ('6821802', 'Corretagem no aluguel de imóveis', False),
        ('6911701', 'Serviços advocatícios', False),
        ('6911703', 'Agente de propriedade industrial', True),
        ('7020400', 'Atividades de consultoria em gestão empresarial, exceto consultoria técnica específica', True),
        ('7111100', 'Serviços de arquitetura', True),
        ('7112000', 'Serviços de engenharia', True),
        ('7119701', 'Serviços de cartografia, topografia e geodésia', True),
        ('7119702', 'Atividades de estudos geológicos', True),
        ('7119703', 'Serviços de desenho técnico relacionados à arquitetura e engenharia', True),
        ('7119704', 'Serviços de perícia técnica relacionados à segurança do trabalho', True),
        ('7119799', 'Atividades técnicas relacionadas à engenharia e arquitetura não especificadas anteriormente', True),
        ('7120100', 'Testes e análises técnicas', True),
        ('7210000', 'Pesquisa e desenvolvimento experimental em ciências físicas e naturais', True),
        ('7220700', 'Pesquisa e desenvolvimento experimental em ciências sociais e humanas', True),
        ('7311400', 'Agências de publicidade', True),
        ('7312200', 'Agenciamento de espaços para publicidade, exceto em veículos de comunicação', False),
        ('7319001', 'Criação de estandes para feiras e exposições', True),
        ('7319002', 'Promoção de vendas', False),
        ('7319003', 'Marketing direto', False),
        ('7319004', 'Consultoria em publicidade', True),
        ('7319099', 'Outras atividades de publicidade não especificadas anteriormente', False),
        ('7320300', 'Pesquisas de mercado e de opinião pública', True),
        ('7410202', 'Design de interiores', True),
        ('7410203', 'Design de produto', True),
        ('7410299', 'Atividades de design não especificadas anteriormente', True),
        ('7420001', 'Atividades de produção de fotografias, exceto aérea e submarina', False),
        ('7420002', 'Atividades de produção de fotografias aéreas e submarinas', False),
        ('7420003', 'Laboratórios fotográficos', False),
        ('7420004', 'Filmagem de festas e eventos', False),
        ('7420005', 'Serviços de microfilmagem', False),
        ('7490101', 'Serviços de tradução, interpretação e similares', True),
        ('7490102', 'Escafandria e mergulho', False),
        ('7490103', 'Serviços de agronomia e de consultoria às atividades agrícolas e pecuárias', True),
        ('7490104', 'Atividades de intermediação e agenciamento de serviços e negócios em geral, exceto imobiliários', True),
        ('7490105', 'Agenciamento de profissionais para atividades esportivas, culturais e artísticas', True),
        ('7490199', 'Outras atividades profissionais, científicas e técnicas não especificadas anteriormente', True),
        ('7500100', 'Atividades veterinárias', True),
        ('7721700', 'Aluguel de equipamentos recreativos e esportivos', False),
        ('7722500', 'Aluguel de fitas de vídeo, DVDs e similares', False),
        ('7723300', 'Aluguel de objetos do vestuário, jóias e acessórios', False),
        ('7729201', 'Aluguel de aparelhos de jogos eletrônicos', False),
        ('7729202', 'Aluguel de móveis, utensílios e aparelhos de uso doméstico e pessoal; instrumentos musicais', False),
        ('7729203', 'Aluguel de material médico', False),
        ('7729299', 'Aluguel de outros objetos pessoais e domésticos não especificados anteriormente', False),
        ('7731400', 'Aluguel de máquinas e equipamentos agrícolas sem operador', False),
        ('7732201', 'Aluguel de máquinas e equipamentos para construção sem operador, exceto andaimes', False),
        ('7732202', 'Aluguel de andaimes', False),
        ('7733100', 'Aluguel de máquinas e equipamentos para escritório', False),
        ('7739001', 'Aluguel de máquinas e equipamentos para extração de minérios e petróleo, sem operador', False),
        ('7739002', 'Aluguel de equipamentos científicos, médicos e hospitalares, sem operador', False),
        ('7739003', 'Aluguel de palcos, coberturas e outras estruturas de uso temporário, exceto andaimes', False),
        ('7739099', 'Aluguel de outras máquinas e equipamentos comerciais e industriais não especificados anteriormente, sem operador', False),
        ('7740300', 'Gestão de ativos intangíveis não-financeiros', True),
        ('7911200', 'Agências de viagens', False),
        ('7912100', 'Operadores turísticos', False),
        ('7990200', 'Serviços de reservas e outros serviços de turismo não especificados anteriormente', False),
        ('8011102', 'Serviços de adestramento de cães de guarda', False),
        ('8030700', 'Atividades de investigação particular', True),
        ('8121400', 'Limpeza em prédios e em domicílios', False),
        ('8122200', 'Imunização e controle de pragas urbanas', False),
        ('8129000', 'Atividades de limpeza não especificadas anteriormente', False),
        ('8130300', 'Atividades paisagísticas', False),
        ('8211300', 'Serviços combinados de escritório e apoio administrativo', False),
        ('8219999', 'Preparação de documentos e serviços especializados de apoio administrativo não especificados anteriormente', False),
        ('8220200', 'Atividades de teleatendimento', False),
        ('8230001', 'Serviços de organização de feiras, congressos, exposições e festas', False),
        ('8230002', 'Casas de festas e eventos', False),
        ('8291100', 'Atividades de cobrança e informações cadastrais', False),
        ('8299701', 'Medição de consumo de energia elétrica, gás e água', False),
        ('8299703', 'Serviços de gravação de carimbos, exceto confecção', False),
        ('8299707', 'Salas de acesso à internet', False),
        ('8299799', 'Outras atividades de serviços prestados principalmente às empresas não especificadas anteriormente', True),
        ('8511200', 'Educação infantil - creche', False),
        ('8512100', 'Educação infantil - pré-escola', False),
        ('8513900', 'Ensino fundamental', False),
        ('8520100', 'Ensino médio', False),
        ('8531700', 'Educação superior - graduação', True),
        ('8532500', 'Educação superior - graduação e pós-graduação', True),
        ('8533300', 'Educação superior - pós-graduação e extensão', True),
        ('8541400', 'Educação profissional de nível técnico', False),
        ('8542200', 'Educação profissional de nível tecnológico', True),
        ('8550302', 'Atividades de apoio à educação, exceto caixas escolares', True),
        ('8591100', 'Ensino de esportes', True),
        ('8592901', 'Ensino de dança', True),
        ('8592902', 'Ensino de artes cênicas, exceto dança', False),
        ('8592903', 'Ensino de música', False),
        ('8592999', 'Ensino de arte e cultura não especificado anteriormente', False),
        ('8593700', 'Ensino de idiomas', False),
        ('8599603', 'Treinamento em informática', False),
        ('8599604', 'Treinamento em desenvolvimento profissional e gerencial', False),
        ('8599605', 'Cursos preparatórios para concursos', False),
        ('8599699', 'Outras atividades de ensino não especificadas anteriormente', False),
        ('8610101', 'Atividades de atendimento hospitalar, exceto pronto-socorro e unidades para atendimento a urgências', True),
        ('8610102', 'Atividades de atendimento em pronto-socorro e unidades hospitalares para atendimento a urgências', True),
        ('8621601', 'UTI móvel', True),
        ('8621602', 'Serviços móveis de atendimento a urgências, exceto por UTI móvel', True),
        ('8622400', 'Serviços de remoção de pacientes, exceto os serviços móveis de atendimento a urgências', True),
        ('8630501', 'Atividade médica ambulatorial com recursos para realização de procedimentos cirúrgicos', True),
        ('8630502', 'Atividade médica ambulatorial com recursos para realização de exames complementares', True),
        ('8630503', 'Atividade médica ambulatorial restrita a consultas', True),
        ('8630504', 'Atividade odontológica', True),
        ('8630506', 'Serviços de vacinação e imunização humana', True),
        ('8630507', 'Atividades de reprodução humana assistida', True),
        ('8630599', 'Atividades de atenção ambulatorial não especificadas anteriormente', True),
        ('8640201', 'Laboratórios de anatomia patológica e citológica', True),
        ('8640202', 'Laboratórios clínicos', True),
        ('8640203', 'Serviços de diálise e nefrologia', True),
        ('8640204', 'Serviços de tomografia', True),
        ('8640205', 'Serviços de diagnóstico por imagem com uso de radiação ionizante, exceto tomografia', True),
        ('8640206', 'Serviços de ressonância magnética', True),
        ('8640207', 'Serviços de diagnóstico por imagem sem uso de radiação ionizante, exceto ressonância magnética', True),
        ('8640208', 'Serviços de diagnóstico por registro gráfico - ECG, EEG e outros exames análogos', True),
        ('8640209', 'Serviços de diagnóstico por métodos ópticos - endoscopia e outros exames análogos', True),
        ('8640210', 'Serviços de quimioterapia', True),
        ('8640211', 'Serviços de radioterapia', True),
        ('8640212', 'Serviços de hemoterapia', True),
        ('8640213', 'Serviços de litotripsia', True),
        ('8640214', 'Serviços de bancos de células e tecidos humanos', True),
        ('8640299', 'Atividades de serviços de complementação diagnóstica e terapêutica não especificadas anteriormente', True),
        ('8650001', 'Atividades de enfermagem', True),
        ('8650002', 'Atividades de profissionais da nutrição', True),
        ('8650003', 'Atividades de psicologia e psicanálise', True),
        ('8650004', 'Atividades de fisioterapia', True),
        ('8650005', 'Atividades de terapia ocupacional', True),
        ('8650006', 'Atividades de fonoaudiologia', True),
        ('8650007', 'Atividades de terapia de nutrição enteral e parenteral', True),
        ('8650099', 'Atividades de profissionais da área de saúde não especificadas anteriormente', True),
        ('8660700', 'Atividades de apoio à gestão de saúde', True),
        ('8690901', 'Atividades de práticas integrativas e complementares em saúde humana', True),
        ('8690902', 'Atividades de banco de leite humano', True),
        ('8690903', 'Atividades de acupuntura', True),
        ('8690904', 'Atividades de podologia', True),
        ('8690999', 'Outras atividades de atenção à saúde humana não especificadas anteriormente', True),
        ('8711501', 'Clínicas e residências geriátricas', True),
        ('8711502', 'Instituições de longa permanência para idosos', False),
        ('8712300', 'Atividades de fornecimento de infra-estrutura de apoio e assistência a paciente no domicílio', False),
        ('9001901', 'Produção teatral', False),
        ('9001902', 'Produção musical', False),
        ('9001903', 'Produção de espetáculos de dança', False),
        ('9001904', 'Produção de espetáculos circenses, de marionetes e similares', False),
        ('9001905', 'Produção de espetáculos de rodeios, vaquejadas e similares', False),
        ('9001906', 'Atividades de sonorização e de iluminação', False),
        ('9001999', 'Artes cênicas, espetáculos e atividades complementares não especificados anteriormente', False),
        ('9002701', 'Atividades de artistas plásticos, jornalistas independentes e escritores', True),
        ('9002702', 'Restauração de obras de arte', False),
        ('9003500', 'Gestão de espaços para artes cênicas, espetáculos e outras atividades artísticas', False),
        ('9101500', 'Atividades de bibliotecas e arquivos', False),
        ('9102301', 'Atividades de museus e de exploração de lugares e prédios históricos e atrações similares', False),
        ('9102302', 'Restauração e conservação de lugares e prédios históricos', False),
        ('9311500', 'Gestão de instalações de esportes', False),
        ('9313100', 'Atividades de condicionamento físico', True),
        ('9319101', 'Produção e promoção de eventos esportivos', False),
        ('9319199', 'Outras atividades esportivas não especificadas anteriormente', False),
        ('9329802', 'Exploração de boliches', False),
        ('9329803', 'Exploração de jogos de sinuca, bilhar e similares', False),
        ('9329804', 'Exploração de jogos eletrônicos recreativos', False),
        ('9329899', 'Outras atividades de recreação e lazer não especificadas anteriormente', False),
        ('9511800', 'Reparação e manutenção de computadores e de equipamentos periféricos', False),
        ('9512600', 'Reparação e manutenção de equipamentos de comunicação', False),
        ('9521500', 'Reparação e manutenção de equipamentos eletroeletrônicos de uso pessoal e doméstico', False),
        ('9529101', 'Reparação de calçados, bolsas e artigos de viagem', False),
        ('9529102', 'Chaveiros', False),
        ('9529103', 'Reparação de relógios', False),
        ('9529104', 'Reparação de bicicletas, triciclos e outros veículos não-motorizados', False),
        ('9529105', 'Reparação de artigos do mobiliário', False),
        ('9529106', 'Reparação de jóias', False),
        ('9529199', 'Reparação e manutenção de outros objetos e equipamentos pessoais e domésticos não especificados anteriormente', False),
        ('9601701', 'Lavanderias', False),
        ('9601702', 'Tinturarias', False),
        ('9601703', 'Toalheiros', False),
        ('9602501', 'Cabeleireiros, manicure e pedicure', False),
        ('9602502', 'Atividades de Estética e outros serviços de cuidados com a beleza', False),
        ('9609202', 'Agências matrimoniais', False),
        ('9609205', 'Atividades de sauna e banhos', False),
        ('9609206', 'Serviços de tatuagem e colocação de piercing', False),
        ('9609207', 'Alojamento de animais domésticos', False),
        ('9609208', 'Higiene e embelezamento de animais domésticos', False),
        ('9609299', 'Outras atividades de serviços pessoais não especificadas anteriormente', False),
        ('161003', 'Serviço de preparação de terreno, cultivo e colheita', False),
        ('162803', 'Serviço de manejo de animais', False),
        ('322107', 'Atividades de apoio à aquicultura em água doce', False),
        ('1099604', 'Fabricação de gelo comum', False),
        ('3312102', 'Manutenção e reparação de aparelhos e instrumentos de medida, teste e controle', False),
        ('3312103', 'Manutenção e reparação de aparelhos eletromédicos e eletroterapêuticos e equipamentos de irradiação', False),
        ('3312104', 'Manutenção e reparação de equipamentos e instrumentos ópticos', False),
        ('3313901', 'Manutenção e reparação de geradores, transformadores e motores elétricos', False),
        ('3313902', 'Manutenção e reparação de baterias e acumuladores elétricos, exceto para veículos', False),
        ('3313999', 'Manutenção e reparação de máquinas, aparelhos e materiais elétricos não especificados anteriormente', False),
        ('3314701', 'Manutenção e reparação de máquinas motrizes não-elétricas', False),
        ('3314702', 'Manutenção e reparação de equipamentos hidráulicos e pneumáticos, exceto válvulas', False),
        ('3314703', 'Manutenção e reparação de válvulas industriais', False),
        ('3314704', 'Manutenção e reparação de compressores', False),
        ('3314705', 'Manutenção e reparação de equipamentos de transmissão para fins industriais', False),
        ('3314706', 'Manutenção e reparação de máquinas, aparelhos e equipamentos para instalações térmicas', False),
        ('3314707', 'Manutenção e reparação de máquinas e aparelhos de refrigeração e ventilação para uso industrial e comercial', False),
        ('3314708', 'Manutenção e reparação de máquinas, equipamentos e aparelhos para transporte e elevação de cargas', False),
        ('3314709', 'Manutenção e reparação de máquinas de escrever, calcular e de outros equipamentos não-eletrônicos para escritório', False),
        ('3314710', 'Manutenção e reparação de máquinas e equipamentos para uso geral não especificados anteriormente', False),
        ('3314711', 'Manutenção e reparação de máquinas e equipamentos para agricultura e pecuária', False),
        ('3314712', 'Manutenção e reparação de tratores agrícolas', False),
        ('3314713', 'Manutenção e reparação de máquinas-ferramenta', False),
        ('3314714', 'Manutenção e reparação de máquinas e equipamentos para a prospecção e extração de petróleo', False),
        ('3314715', 'Manutenção e reparação de máquinas e equipamentos para uso na extração mineral, exceto na extração de petróleo', False),
        ('3314716', 'Manutenção e reparação de tratores, exceto agrícolas', False),
        ('3314717', 'Manutenção e reparação de máquinas e equipamentos de terraplenagem, pavimentação e construção, exceto tratores', False),
        ('3314718', 'Manutenção e reparação de máquinas para a indústria metalúrgica, exceto máquinas-ferramenta', False),
        ('3314719', 'Manutenção e reparação de máquinas e equipamentos para as indústrias de alimentos, bebidas e fumo', False),
        ('3314720', 'Manutenção e reparação de máquinas e equipamentos para a indústria têxtil, do vestuário, do couro e calçados', False),
        ('3314721', 'Manutenção e reparação de máquinas e aparelhos para a indústria de celulose, papel e papelão e artefatos', False),
        ('3314722', 'Manutenção e reparação de máquinas e aparelhos para a indústria do plástico', False),
        ('3314799', 'Manutenção e reparação de outras máquinas e equipamentos para usos industriais não especificados anteriormente', False),
        ('3315500', 'Manutenção e reparação de veículos ferroviários', False),
        ('3316301', 'Manutenção e reparação de aeronaves, exceto a manutenção na pista', False),
        ('3316302', 'Manutenção de aeronaves na pista', False),
        ('3317101', 'Manutenção e reparação de embarcações e estruturas flutuantes', False),
        ('3317102', 'Manutenção e reparação de embarcações para esporte e lazer', False),
        ('3319800', 'Manutenção e reparação de equipamentos e produtos não especificados anteriormente', False),
        ('3321000', 'Instalação de máquinas e equipamentos industriais', False),
        ('4211102', 'Pintura para sinalização em pistas rodoviárias e aeroportos', False),
        ('4221903', 'Manutenção de redes de distribuição de energia elétrica', False),
        ('4221905', 'Manutenção de estações e redes de telecomunicações', False),
        ('4313400', 'Obras de terraplenagem', False),
        ('4321500', 'Instalação e manutenção elétrica', False),
        ('4322301', 'Instalações hidráulicas, sanitárias e de gás', False),
        ('4322302', 'Instalação e manutenção de sistemas centrais de ar condicionado, de ventilação e refrigeração', False),
        ('4322303', 'Instalações de sistema de prevenção contra incêndio', False),
        ('4329101', 'Instalação de painéis publicitários', False),
        ('4329102', 'Instalação de equipamentos para orientação à navegação marítima, fluvial e lacustre', False),
        ('4329103', 'Instalação, manutenção e reparação de elevadores, escadas e esteiras rolantes', False),
        ('4329104', 'Montagem e instalação de sistemas e equipamentos de iluminação e sinalização em vias públicas, portos e aeroportos', False),
        ('4329105', 'Tratamentos térmicos, acústicos ou de vibração', False),
        ('4399102', 'Montagem e desmontagem de andaimes e outras estruturas temporárias', False),
        ('4511101', 'Comércio a varejo de automóveis, camionetas e utilitários novos', False),
        ('4520001', 'Serviços de manutenção e reparação mecânica de veículos automotores', False),
        ('4520002', 'Serviços de lanternagem ou funilaria e pintura de veículos automotores', False),
        ('4520003', 'Serviços de manutenção e reparação elétrica de veículos automotores', False),
        ('4520004', 'Serviços de alinhamento e balanceamento de veículos automotores', False),
        ('4520005', 'Serviços de lavagem, lubrificação e polimento de veículos automotores', False),
        ('4520006', 'Serviços de borracharia para veículos automotores', False),
        ('4520007', 'Serviços de instalação, manutenção e reparação de acessórios para veículos automotores', False),
        ('4520008', 'Serviços de capotaria', False),
        ('4530701', 'Comércio por atacado de peças e acessórios novos para veículos automotores', False),
        ('4541203', 'Comércio a varejo de motocicletas e motonetas novas', False),
        ('4541204', 'Comércio a varejo de motocicletas e motonetas usadas', False),
        ('4543900', 'Manutenção e reparação de motocicletas e motonetas', False),
        ('4722902', 'Peixaria', False),
        ('4729602', 'Comércio varejista de mercadorias em lojas de conveniência', False),
        ('4741500', 'Comércio varejista de tintas e materiais para pintura', False),
        ('4742300', 'Comércio varejista de material elétrico', False),
        ('4744001', 'Comércio varejista de ferragens e ferramentas', False),
        ('4744003', 'Comércio varejista de materiais hidráulicos', False),
        ('4744004', 'Comércio varejista de cal, areia, pedra britada, tijolos e telhas', False),
        ('4744005', 'Comércio varejista de materiais de construção não especificados anteriormente', False),
        ('4744006', 'Comércio varejista de pedras para revestimento', False),
        ('4744099', 'Comércio varejista de materiais de construção em geral', False),
        ('4763605', 'Comércio varejista de embarcações e outros veículos recreativos; peças e acessórios', False),
        ('4771704', 'Comércio varejista de medicamentos veterinários', False),
        ('4773300', 'Comércio varejista de artigos médicos e ortopédicos', False),
        ('4789005', 'Comércio varejista de produtos saneantes domissanitários', False),
        ('4789006', 'Comércio varejista de fogos de artifício e artigos pirotécnicos', False),
        ('4789009', 'Comércio varejista de armas e munições', False),
        ('5211702', 'Guarda-móveis', False),
        ('5320201', 'Serviços de malote não realizados pelo Correio Nacional', False),
        ('5320202', 'Serviços de entrega rápida', False),
        ('5510801', 'Hotéis', False),
        ('5510802', 'Apart-hotéis', False),
        ('5510803', 'Motéis', False),
        ('5590601', 'Albergues, exceto assistenciais', False),
        ('5590602', 'Campings', False),
        ('5590603', 'Pensões (alojamento)', False),
        ('5590699', 'Outros alojamentos não especificados anteriormente', False),
        ('5620102', 'Serviços de alimentação para eventos e recepções - bufê', False),
        ('5811500', 'Edição de livros', False),
        ('5812301', 'Edição de jornais diários', False),
        ('5812302', 'Edição de jornais não diários', False),
        ('5813100', 'Edição de revistas', False),
        ('5819100', 'Edição de cadastros, listas e outros produtos gráficos', False),
        ('6822600', 'Gestão e administração da propriedade imobiliária', True),
        ('7711000', 'Locação de automóveis sem condutor', False),
        ('8020001', 'Atividades de monitoramento de sistemas de segurança eletrônico', False),
        ('8219901', 'Fotocópias', False),
        ('8599601', 'Formação de condutores', False),
        ('8599602', 'Cursos de pilotagem', False),
        ('6612605', 'Agentes de investimentos em aplicações financeiras', False),
        ('6911702', 'Atividades auxiliares da justiça', False),
    ]
    try:
        cursor.executemany('INSERT OR REPLACE INTO cnae_fator_r VALUES (?, ?, ?)', cnae_data)
        conn.commit()
    except sqlite3.Error as e:
        print(f"Erro ao popular o banco de dados: {e}")
        conn.close()
        exit(1)

# Chame esta função uma vez para popular o banco de dados
populate_database()

# ... (mantenha as funções de banco de dados e populate_database)

def is_fator_r(cnae_code):
    try:
        # Extrair apenas o código numérico do CNAE
        cnae_code = cnae_code.split(' - ')[0].strip()
        cursor.execute('SELECT fator_r FROM cnae_fator_r WHERE cnae = ?', (cnae_code,))
        result = cursor.fetchone()
        if result:
            return result[0]
        return False
    except sqlite3.Error as e:
        print(f"Erro ao consultar o banco de dados: {e}")
        return False

def validate_cnpj(cnpj):
    # Remove caracteres não numéricos
    cnpj = re.sub(r'\D', '', cnpj)
    
    # Verifica se tem 14 dígitos
    if len(cnpj) != 14:
        return False
    
    # Verifica se todos os dígitos são iguais
    if len(set(cnpj)) == 1:
        return False
    
    # Implementação da validação do dígito verificador pode ser adicionada aqui
    
    return True

def extract_cnae_info(soup):
    cnaes = []
    
    # Encontrar a tabela de atividades econômicas usando o novo seletor
    table = soup.select_one("#page-content > div > div.flex-1.overflow-x-auto.min-h-screen.px-5.pb-40.pt-8.md\:px-10 > div:nth-child(10) > div > table")
    
    if table:
        rows = table.find_all('tr')[1:]  # Ignorar o cabeçalho
        for row in rows:
            cols = row.find_all('td')
            if len(cols) == 3:
                tipo = 'Principal' if cols[0].find('span').text.strip() == 'P' else 'Secundário'
                codigo = cols[1].find('span').text.strip()
                # Remover "-" e "/" do código CNAE
                codigo_limpo = codigo.replace('-', '').replace('/', '')
                descricao = cols[2].find('span').text.strip()
                cnaes.append((tipo, f"{codigo_limpo} - {descricao}"))
    
    print("CNAEs extraídos:", cnaes)
    return cnaes

def fill_cnpj_on_site(cnpj, progress_var):
    url = "https://cnpja.com/office/23890519000198"

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_experimental_option("useAutomationExtension", False)
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        print(f"Erro ao inicializar o driver: {e}")
        messagebox.showerror("Erro", "Não foi possível inicializar o driver do Chrome. Verifique se o Chrome está instalado.")
        return None

    try:
        driver.get(url)
        progress_var.set(20)  # 20% progress
        
        cnpj_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="page-content"]/div/div[1]/label/div/input'))
        )
        cnpj_input.clear()
        
        for digit in cnpj:
            cnpj_input.send_keys(digit)
            time.sleep(0.1)
        
        driver.execute_script("arguments[0].value = arguments[1]", cnpj_input, cnpj)
        progress_var.set(40)  # 40% progress

        time.sleep(7)

        # Aguardar o carregamento dos dados
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#page-content > div > div.flex-1.overflow-x-auto.min-h-screen.px-5.pb-40.pt-8.md\:px-10 > div:nth-child(10) > div > table"))
        )

        content = driver.page_source
        soup = BeautifulSoup(content, 'html.parser')
        
        cnaes = extract_cnae_info(soup)
        
        if not cnaes:
            print("Nenhum dado CNAE encontrado")
            return None
        
        progress_var.set(100)  # 100% progress
        return cnaes

    except Exception as e:
        print(f"Ocorreu um erro ao extrair os dados: {e}")
        return None

    finally:
        driver.quit()


def clean_non_digits():
    current_text = cnpj_entry.get()
    cleaned_text = ''.join(filter(str.isdigit, current_text))
    if current_text != cleaned_text:
        cnpj_entry.delete(0, tk.END)
        cnpj_entry.insert(0, cleaned_text)

def search_cnpj():
    cnpj = cnpj_entry.get()
    if not cnpj:
        messagebox.showerror("Erro", "Por favor, insira um CNPJ")
        return

    if not validate_cnpj(cnpj):
        messagebox.showerror("Erro", "CNPJ inválido. Por favor, verifique o número inserido.")
        return

    progress_var.set(0)
    progress_bar.pack(fill=tk.X, pady=(10, 0))
    loading_label.pack(pady=(5, 0))
    search_button.config(state=tk.DISABLED)
    root.update_idletasks()

    def process_search():
        nonlocal cnpj
        result = fill_cnpj_on_site(cnpj, progress_var)
        root.after(0, update_ui, result)

    threading.Thread(target=process_search, daemon=True).start()

def update_ui(result):
    if result:
        tree.delete(*tree.get_children())

        for tipo, cnae_info in result:
            try:
                cnae_code, cnae_description = cnae_info.split(' - ', 1)
                full_cnae = f"{cnae_code} - {cnae_description}"
                fator_r = is_fator_r(cnae_code)
                tree.insert('', 'end', values=(tipo, full_cnae, "SIM" if fator_r else "NÃO"))
            except ValueError as e:
                print(f"Erro ao processar item CNAE: {e}")
                print(f"Item problemático: {cnae_info}")
    else:
        messagebox.showerror("Erro", "Falha ao recuperar os dados")

    progress_bar.pack_forget()
    loading_label.pack_forget()
    search_button.config(state=tk.NORMAL)

# Configuração da GUI
root = tk.Tk()
root.title("Consulta de CNPJ")
root.geometry("800x600")  # Define um tamanho inicial para a janela

frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

cnpj_frame = ttk.Frame(frame)
cnpj_frame.pack(fill=tk.X, pady=(0, 10))

cnpj_label = ttk.Label(cnpj_frame, text="CNPJ:")
cnpj_label.pack(side=tk.LEFT, padx=(0, 5))

cnpj_entry = ttk.Entry(cnpj_frame, width=30)
cnpj_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
cnpj_entry.bind('<KeyRelease>', lambda e: clean_non_digits())

search_button = ttk.Button(cnpj_frame, text="Pesquisar", command=search_cnpj)
search_button.pack(side=tk.LEFT)

tree = ttk.Treeview(frame, columns=('Tipo', 'CNAE', 'Fator R'), show='headings')
tree.heading('Tipo', text='Tipo')
tree.heading('CNAE', text='CNAE - Descrição')
tree.heading('Fator R', text='Fator R')
tree.column('Tipo', width=100)
tree.column('CNAE', width=500)
tree.column('Fator R', width=100)
tree.pack(fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
tree.configure(yscrollcommand=scrollbar.set)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(frame, variable=progress_var, maximum=100)

loading_label = ttk.Label(frame, text="Carregando... Por favor, aguarde.")

if __name__ == "__main__":
    populate_database()
    root.mainloop()
    conn.close()


# Fechar a conexão com o banco de dados ao sair do programa
#conn.close()