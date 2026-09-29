# Comparador Ergon x Intranet

Aplicativo Streamlit para comparar servidores da folha de pagamento do **Ergon** com a escala cadastrada na **Intranet**. O objetivo é localizar servidores que estão na folha, mas não aparecem escalados no hospital correspondente.

## Regra de comparação

A chave usada nas duas planilhas é composta por `NUMFUNC` e `NUMVINC`, no formato `NUMFUNC-NUMVINC`. O CPF não é usado como chave principal. Se o mesmo vínculo aparecer em várias linhas da Intranet por possuir mais de uma escala, basta uma ocorrência no hospital correto para que ele seja considerado escalado.

O campo `LOTACAO` do Ergon é relacionado ao campo `SETOR` da Intranet pelo mapeamento dos hospitais informado no projeto. Lotações do Ergon que não estejam mapeadas aparecem em uma aba separada e não entram no cálculo de não escalados.

## Como executar

```bash
python3 -m pip install -r requirements.txt
streamlit run app.py
```

Envie uma planilha do Ergon e uma da Intranet. O Ergon deve conter pelo menos `NUMFUNC`, `NUMVINC` e `LOTACAO`; a Intranet deve conter `NUMFUNC`, `NUMVINC` e `SETOR`. O sistema aceita CSV, XLS, XLSX e XLSM e reconhece pequenas variações nos cabeçalhos.

## Saída

A tela exibe indicadores, lista filtrável por hospital, resumo por hospital e lotações fora do mapeamento. O Excel exportado respeita o filtro escolhido e é configurado para impressão em **A4 paisagem**, com ajuste para **uma página de largura**, cabeçalho repetido e área de impressão.

## Testes

```bash
pytest -q
python3 -m py_compile app.py test_app.py
```
