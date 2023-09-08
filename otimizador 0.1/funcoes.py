"""Funções comuns aos diferentes modelos implementados"""

import numpy as np

def imprimir_resultados(qtdes, n_perfis, n_unidades, matriz_unidades, nomes_unidades, \
                        matriz_peq, matriz_tempo):
    """Imprime resultados da quantidade de cada perfil em cada unidade"""
    relatorio = "\nResultados:"
    borda_cabecalho = "\n---------+" + "-----"*n_perfis \
        + "-+-------+---------+----------+------------+"
    # Cabeçalho
    relatorio += borda_cabecalho
    relatorio += "\nUnidade  |  " \
        + "  ".join([f"{i: >3}" for i in [f"p{p+1}" for p in range(n_perfis)]]) \
        + " | Total |   P-Eq  |   Tempo  | Tempo/prof |"
    relatorio += borda_cabecalho
    # Uma linha por unidade
    for unidade in range(n_unidades):
        total = np.sum(qtdes[unidade])
        peq = np.sum(qtdes[unidade]*matriz_peq)
        tempo = np.sum(qtdes[unidade]*matriz_tempo) - matriz_unidades[unidade][1]
        relatorio += f"\n{nomes_unidades[unidade]:6s}   | " \
            + " ".join([f"{qtdes[unidade][p]:4d}" for p in range(n_perfis)]) \
            + f" |  {total:4d} | {peq:7.2f} |  {tempo:7.2f} |    {(tempo)/total:7.3f} |"

    # Totais
    total = np.sum(qtdes)
    peq = np.sum(qtdes*matriz_peq)
    tempo = np.sum(qtdes*matriz_tempo) - np.sum(matriz_unidades[:n_unidades], axis=0)[1]

    relatorio += borda_cabecalho
    relatorio += "\nTotal    | " \
        + " ".join([f"{np.sum(qtdes, axis=0)[p]:4d}" for p in range(n_perfis)])
    relatorio += f" |  {total:4d} | {peq:7.2f} |  {tempo:7.2f} |    {(tempo)/total:7.3f} |"
    relatorio += borda_cabecalho + "\n"

    return relatorio


def imprimir_parametros(qtdes, n_unidades, n_restricoes, matriz_unidades, nomes_unidades, \
                        restricoes_percentuais, matriz_perfis, nomes_restricoes):
    """Imprime os dados de entrada e os resultados obtidos"""
    relatorio = "\nParâmetros:"
    # Formatos dos números - tem que ser tudo como float, pois ao importar os valores de
    # professor-equivalente, a matriz_perfis fica toda como float
    formatos =      ['4.0f', '7.2f', '4.0f', '3.0f', '3.0f']
    formados_dif =  ['3.0f', '7.2f', '4.0f', '2.0f', '2.0f']
    formatos_nome = [11, 18, 12, 9, 9]

    borda_cabecalho = "\n---------+"
    linha_cabecalho = "\nUnidade  |"

    # Para cada restrição, coloca o nome na linha do cabeçalho
    # e na borda é preciso adicionar dois traços
    for idx, nome in enumerate(nomes_restricoes):
        borda_cabecalho += "-"*(formatos_nome[idx]+2) + "+"
        linha_cabecalho += f" {nome:>{formatos_nome[idx]}} |"

    borda_cabecalho += "-----------+"
    linha_cabecalho += "  ch media |"

    # Se houver restrições em algum perfil, acrescenta colunas
    if len(restricoes_percentuais) > 0:
        for restricao in restricoes_percentuais.values():
            # Acrescenta ao cabeçalho
            perfis = restricao['perfis']
            texto = 'P ' + ','.join(str(p+1) for p in perfis)
            texto = texto.center(11) + '|'
            borda_cabecalho += "-----------+"
            linha_cabecalho += texto

    relatorio += borda_cabecalho
    relatorio += linha_cabecalho
    relatorio += borda_cabecalho

    # Uma linha por unidade
    for unidade in range(n_unidades):
        total_unidade = np.sum(qtdes[unidade])
        valores_perfis = [np.sum(qtdes[unidade]*matriz_perfis[p]) for p in range(n_restricoes)]
        diferencas = [valores_perfis[p] - matriz_unidades[unidade][p]
                      for p in range(n_restricoes)]
        strings_perfis = [f"{valores_perfis[p]:{formatos[p]}} " \
                          f"(+{diferencas[p]:{formados_dif[p]}}) |"
                            for p in range(n_restricoes)]
        string_final = " ".join(strings_perfis)

        relatorio += f"\n{nomes_unidades[unidade]:6s}   | " + string_final \
            + f"  {matriz_unidades[unidade][0] / total_unidade:8.4f} |"
        # Se houver restrições em algum perfil, acrescenta colunas
        if len(restricoes_percentuais) > 0:
            for restricao in restricoes_percentuais.values():
                perfis = restricao['perfis']
                # Calcula percentuais
                quantidade = sum(qtdes[unidade][perfil] for perfil in perfis)
                perc_unidade = quantidade / total_unidade * 100
                relatorio += f"   {perc_unidade:6.2f}% |"

    # Totais
    total = np.sum(qtdes)
    valores_perfis = [np.sum(qtdes*matriz_perfis[p]) for p in range(n_restricoes)]
    diferencas = [valores_perfis[p] - int(np.sum(matriz_unidades, axis=0)[p])
                  for p in range(n_restricoes)]
    strings_perfis = [f"{valores_perfis[p]:{formatos[p]}} " \
                      f"(+{diferencas[p]:{formados_dif[p]}}) |" for p in range(n_restricoes)]
    string_final = " ".join(strings_perfis)

    relatorio += borda_cabecalho
    relatorio += "\nTotal    | " + string_final \
        + f"  {np.sum(matriz_unidades, axis=0)[0]/total:8.4f} |"
    # Se houver restrições em algum perfil, acrescenta colunas
    if len(restricoes_percentuais) > 0:
        for restricao in restricoes_percentuais.values():
            perfis = restricao['perfis']
            # Quantidade total
            quantidade_total = np.sum(qtdes[:, perfis])
            perc_total = quantidade_total / total * 100
            relatorio += f"   {perc_total:6.2f}% |"
    relatorio += borda_cabecalho + "\n"

    return relatorio


def imprimir_unidades(n_unidades, n_restricoes, matriz_unidades, nomes_unidades, nomes_restricoes):
    """Imprime os dados de entrada das unidades"""
    relatorio = "\n\nUnidades:"
    #borda_cabecalho = "\n---------+-------+--------------+------------+---------+---------+"
    ##------------------- Usar nomes das restrições aqui ---------------------##
    #linha_cabecalho = "\nUnidade  | aulas | horas_orient | num_orient | diretor | coords. |"
    borda_cabecalho = "\n---------+"
    for restricao in nomes_restricoes:
        borda_cabecalho += "-"*10 + "+"
    linha_cabecalho = "\nUnidade  |"
    linha_cabecalho = "\nUnidade  |" + \
        "".join([f" {restricao:>8} |" for restricao in nomes_restricoes])

    relatorio += borda_cabecalho
    relatorio += linha_cabecalho
    relatorio += borda_cabecalho
    # Formatos dos números - tem que ser tudo como float, pois ao importar os valores de
    # professor-equivalente, a matriz_perfis fica toda como float
    formatos =          ['8.0f', '8.2f', '8.0f', '8.0f', '8.0f']

    # Uma linha por unidade
    for unidade in range(n_unidades):
        valores_restricoes = [matriz_unidades[unidade][p] for p in range(n_restricoes)]
        strings_restricoes = [f"{valores_restricoes[p]:{formatos[p]}} |"
                              for p in range(n_restricoes)]
        string_final = " ".join(strings_restricoes)

        relatorio += f"\n{nomes_unidades[unidade]:6s}   | " + string_final

    # Totais
    valores_perfis = [np.sum(matriz_unidades, axis=0)[p] for p in range(n_restricoes)]
    strings_perfis = [f"{valores_perfis[p]:{formatos[p]}} |" for p in range(n_restricoes)]
    string_final = " ".join(strings_perfis)

    relatorio += borda_cabecalho
    relatorio += "\nTotal    | " + string_final
    relatorio += borda_cabecalho + "\n"

    return relatorio


def imprimir_perfis(n_perfis, n_restricoes, matriz_perfis, nomes_restricoes):
    """Imprime os perfis utilizados"""
    relatorio = '\nPerfis:'
    borda_cabecalho = '\n---------------+' + '------+'*n_perfis
    linha_cabecalho = '\nCaracterística |' \
        + "".join([f"  {i: >3} |" for i in [f"p{p+1}" for p in range(n_perfis)]])
        #+ ''.join([f'  P{p+1}  |' for p in range(n_perfis)])

    relatorio += borda_cabecalho
    relatorio += linha_cabecalho
    relatorio += borda_cabecalho

    for restricao in range(n_restricoes):
        relatorio += f'\n{nomes_restricoes[restricao]:14s} | ' \
            + ' '.join([f'{matriz_perfis[restricao][p]:4.0f} |' for p in range(n_perfis)])

    relatorio += '\nOutras ativ.   | ' \
        + ' '.join([f'{matriz_perfis[n_restricoes][p]:4.0f} |' for p in range(n_perfis)])

    relatorio += '\nProf-equiv.    | ' \
        + ' '.join([f'{matriz_perfis[n_restricoes+1][p]:4.2f} |' for p in range(n_perfis)])

    relatorio += borda_cabecalho + "\n"

    return relatorio
