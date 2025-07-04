SELECT OCO.numero_ocorrencia,   -- Número da ocorrência                                      
OCO.natureza_codigo,                                         -- Código da natureza da ocorrência
OCO.natureza_descricao,                                      -- Descrição da natureza da ocorrência
    CASE
            WHEN OCO.codigo_municipio IN (310090 , 310100 , 310170 , 310270 , 310340 , 310470 , 310520 , 310660 , 311080 , 311300 , 311370 , 311545 , 311700 , 311950 , 312015 , 312235 , 312245 , 312560 , 312675 , 312680 , 312705 , 313230 , 313270 , 313330 , 313400 , 313470 , 313507 , 313580 , 313600 , 313650 , 313700 , 313890 , 313920 , 314055 , 314140 , 314315 , 314430 , 314490 , 314530 , 314535 , 314620 , 314630 , 314675 , 314850 , 314870 , 315000 , 315217 , 315240 , 315510 , 315660 , 315710 , 315765 , 315810 , 316030 , 316330 , 316555 , 316670 , 316860 , 317030 , 317160) THEN '15 RPM'		       
    END AS RPM_2025_AREA,
CASE 
        WHEN OCO.codigo_municipio in (310470,311080,311300,311545,312675,312680,313230,313270,313507,313700,313920,314490,314530,314535,314620,314850,315000,315240,316330,316555,316860) THEN '19 BPM'
        ELSE 'OUTROS'
    END AS UEOP_2025_AREA,
LO.codigo_unidade_area,										-- Código da unidade militar da área
LO.unidade_area_militar_nome,                                -- Nome da unidade militar da área
OCO.unidade_responsavel_registro_codigo,                      -- Código da unidade que registrou a ocorrência
OCO.unidade_responsavel_registro_nome,                        -- Nome da unidade que registrou a ocorrência
SPLIT_PART(OCO.unidade_responsavel_registro_nome,'/',-1) RPM_REGISTRO, 
SPLIT_PART(OCO.unidade_responsavel_registro_nome,'/',-2) UEOP_REGISTRO, 
CASE 																			-- se o território é Urbano ou Rural segundo o IBGE
    	WHEN OCO.pais_codigo <> 1 AND OCO.ocorrencia_uf IS NULL THEN 'Outro_Pais'  	-- trata erro - ocorrencia de fora do Brasil
		WHEN OCO.ocorrencia_uf <> 'MG' THEN 'Outra_UF'								-- trata erro - ocorrencia de fora de MG
    	WHEN OCO.numero_latitude IS NULL THEN 'Invalido'							-- trata erro - ocorrencia sem latitude
        WHEN geo.situacao_codigo = 9 THEN 'Agua'									-- trata erro - ocorrencia dentro de curso d'água
       	WHEN geo.situacao_zona IS NULL THEN 'Erro_Processamento'					-- checa se restou alguma ocorrencia com erro
    	ELSE geo.situacao_zona
END AS situacao_zona,  
CAST(OCO.codigo_municipio AS INTEGER) codigo_municipio,                        -- Converte o código do município para número inteiro
OCO.nome_municipio,                                           -- Nome do município da ocorrência
OCO.tipo_logradouro_descricao,                                -- Tipo do logradouro (Rua, Avenida, etc)
OCO.logradouro_nome,                                          -- Nome do logradouro
OCO.numero_endereco,                                          -- Número do endereço
OCO.nome_bairro,                                              -- Nome do bairro
OCO.ocorrencia_uf,                                            -- Estado da ocorrência
OCO.numero_latitude,                                          -- Latitude da localização
OCO.numero_longitude,                                         -- Longitude da localização
 CONCAT(
    SUBSTR(CAST(OCO.data_hora_fato AS STRING), 9, 2), '/',  -- Dia (posições 9-10)
    SUBSTR(CAST(OCO.data_hora_fato AS STRING), 6, 2), '/',  -- Mês (posições 6-7)
    SUBSTR(CAST(OCO.data_hora_fato AS STRING), 1, 4), ' ',  -- Ano (posições 1-4)
    SUBSTR(CAST(OCO.data_hora_fato AS STRING), 12, 8)       -- Hora (posições 12-19)
  ) AS data_hora_fato2,                   -- Converte a data/hora do fato para o padrão brasileiro
YEAR(OCO.data_hora_fato) AS ano,                           -- Ano do fato
MONTH(OCO.data_hora_fato) AS mes, -- Mês do fato
OCO.data_hora_fato,
OCO.nome_tipo_relatorio,                                   -- Tipo do relatório
OCO.digitador_sigla_orgao,                                  -- Sigla do órgão que registrou
geo.latitude_sirgas2000  AS numero_latitude,                                          -- Latitude da localização
geo.longitude_sirgas2000 AS numero_longitude,                                         -- Longitude da localização
geo.latitude_sirgas2000,				-- reprojeção da latitude de SAD69 para SIRGAS2000
geo.longitude_sirgas2000				-- reprojeção da longitude de SAD69 para SIRGAS2000
FROM db_bisp_reds_reporting.tb_ocorrencia OCO
LEFT JOIN db_bisp_reds_master.tb_local_unidade_area_pmmg LO ON OCO.id_local = LO.id_local
LEFT JOIN db_bisp_reds_master.tb_ocorrencia_setores_geodata AS geo ON OCO.numero_ocorrencia = geo.numero_ocorrencia AND OCO.ocorrencia_uf = 'MG'	-- Tabela de apoio que compara as lat/long com os setores IBGE		
WHERE 1 = 1      -- Condição sempre verdadeira que serve como ponto inicial para facilitar manutenção da query
AND EXISTS (                                                           
-- Predicado que verifica a não existência de registros na subconsulta 
   SELECT 1                                                              
   FROM db_bisp_reds_reporting.tb_envolvido_ocorrencia envolvido          
   WHERE envolvido.numero_ocorrencia = OCO.numero_ocorrencia             
   AND envolvido.id_envolvimento IN(25,1097, 27, 32, 28, 26, 872)        -- Filtro por códigos específicos de envolvimento -  vítimas
   AND envolvido.condicao_fisica_codigo <> '0100'                         -- Filtro por condição física difernte de fatal (código '0100')
)                                                                          -- Fim da subconsulta EXISTS - exclui vítima fatal
AND YEAR(OCO.data_hora_fato) >=2024
--AND OCO.data_hora_fato BETWEEN '2024-01-01 00:00:00.000' AND '2025-04-30 23:59:59.000' -- Filtra ocorrências por período específico (todo o ano de 2024 até fevereiro/2025)
AND OCO.ocorrencia_uf = 'MG'                                                     -- Filtra apenas ocorrências do estado de Minas Gerais
 AND OCO.codigo_municipio IN (310340,311950,313400,314630,317160,311700,313330,314140,315217, -- '70 BPM'
							        310470,311080,311300,311545,312675,312680,313230,313270,313507,313700,313920,314490,314530,314535,314620,314850,315000,315240,316330,316555,316860, -- '19 BPM'
							       310090,310660,311370,312015,312705,313890,314430,315765,316670,317030, -- '24 CIA PM IND'
							       310170,310520,312245,312560,313470,313580,313600,313650,314055,314315,314675,315510,315660,315710,315810,316030,310100,310270,312235,314870) -- '44 BPM'
AND OCO.digitador_sigla_orgao IN ('PM','PC')                              -- Filtro por ocorrências, Polícia Militar ou Polícia Civil
AND OCO.natureza_codigo IN('B01121','B01148','B02001','C01157','C01158','C01159','B01504') -- Seleção de naturezas do CV
AND OCO.ind_estado = 'F'                                                         -- Filtra apenas ocorrências fechadas
--AND OCO.unidade_area_militar_nome LIKE '%15 RPM%'   -- FILTRE PELO NOME DA UNIDADE AREA MILITAR
ORDER BY  OCO.numero_ocorrencia
