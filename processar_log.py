import re
import json


def extrair_pares_limite(caminho_arquivo_log, limite=500):
    """
    Lê um arquivo de log, encontra as extrações que atingiram um limite específico
    e retorna uma lista de dicionários com os detalhes.
    """
    print(
        f"🔍 Analisando o arquivo '{caminho_arquivo_log}' em busca de extrações com limite de {limite} valores...")

    regex = re.compile(
        r"✅\s+([\w\._]+)\s+-\s+([\w\s\+]+?)\s+\(("+str(limite)+r")\s+valores\)")
    pares_com_limite = []

    try:
        with open(caminho_arquivo_log, 'r', encoding='utf-8') as f:
            for linha in f:
                match = regex.search(linha)
                if match:
                    tabela_completa = match.group(1).strip()
                    colunas_str = match.group(2).strip()
                    colunas = [col.strip() for col in colunas_str.split('+')]

                    par = {
                        'tabela': tabela_completa,
                        'coluna_id': colunas[0] if len(colunas) > 0 else None,
                        'coluna_descricao': colunas[1] if len(colunas) > 1 else None
                    }
                    pares_com_limite.append(par)

    except FileNotFoundError:
        print(f"❌ ERRO: O arquivo '{caminho_arquivo_log}' não foi encontrado.")
        return None
    except Exception as e:
        print(f"❌ ERRO: Ocorreu um erro inesperado ao ler o arquivo: {e}")
        return None

    return pares_com_limite


def filtrar_e_otimizar_lista(lista_pares):
    """
    Aplica regras de negócio para remover duplicatas semânticas e descartar campos indesejados.
    """
    print("\n⚙️ Aplicando filtros e otimizações na lista de pares...")

    # --- Palavras-chave para descarte ---
    # Qualquer par que contenha estas palavras em suas colunas será removido.
    palavras_descarte = [
        'logradouro', 'viatura', 'bairro', 'municipio', 'complemento', 'endereco'
    ]

    # --- Chaves únicas para desduplicação ---
    # Usaremos a parte inicial do nome da coluna para identificar duplicatas semânticas.
    # Ex: 'natureza_codigo' e 'natureza_secundaria1_codigo' serão ambos tratados como 'natureza'.
    chaves_unicas_vistas = set()
    lista_otimizada = []

    for par in lista_pares:
        coluna_id = par['coluna_id']

        # 1. Regra de Descarte
        descartar = False
        for palavra in palavras_descarte:
            if palavra in coluna_id:
                descartar = True
                break
        if descartar:
            continue  # Pula para o próximo item da lista

        # 2. Regra de Desduplicação Semântica
        # Pega a primeira parte do nome da coluna como chave (ex: 'natureza' de 'natureza_codigo')
        chave_semantica = coluna_id.split('_')[0]

        # Caso especial para nomes compostos que queremos tratar como um só
        if 'marca_modelo' in coluna_id:
            chave_semantica = 'marca_modelo'

        if chave_semantica not in chaves_unicas_vistas:
            lista_otimizada.append(par)
            chaves_unicas_vistas.add(chave_semantica)

    return lista_otimizada


# --- Execução Principal ---
if __name__ == "__main__":
    arquivo_log = 'log_extracao.txt'
    pares_brutos = extrair_pares_limite(arquivo_log, limite=500)

    if pares_brutos:
        print(
            f"\n✅ Extração inicial concluída! Foram encontrados {len(pares_brutos)} pares que atingiram o limite de 500.")

        # Aplica as regras de negócio para limpar a lista
        pares_finais = filtrar_e_otimizar_lista(pares_brutos)

        print(
            f"✅ Otimização concluída! A lista foi reduzida para {len(pares_finais)} pares únicos e relevantes.")

        # Imprime a lista final de forma legível no terminal
        print("\n--- Lista Final de Pares para Re-extração (Limite 3000) ---")
        for item in pares_finais:
            print(
                f"Tabela: {item['tabela']}, Colunas: {item['coluna_id']} / {item['coluna_descricao']}")

        # Salva a lista final em um arquivo JSON
        arquivo_saida_json = 'pares_finais_para_reextracao.json'
        with open(arquivo_saida_json, 'w', encoding='utf-8') as f_out:
            json.dump(pares_finais, f_out, indent=4, ensure_ascii=False)

        print(
            f"\n💾 A lista final e otimizada foi salva no arquivo: '{arquivo_saida_json}'")
