import os
import modLeitura
import modTratamentoExcel

def main():
    listaOFX = modLeitura.retornarListarArqsOFX('/Volumes/APFS10/FINANCEIRO/OFXs/')
    modLeitura.inserirDadosCSV(listaOFX)
    tabelaFluxoCaixa = modLeitura.retornarListaOperCSV('dados.csv')

    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # Pega a pasta do projeto
    # print(ROOT_DIR)
    arquivo = ROOT_DIR + '/' + 'controle.xlsx'
    modLeitura.inserirTabelaExcel(tabelaFluxoCaixa, arquivo , "Fluxo_Caixa")


if __name__ == "__main__":
    main()
