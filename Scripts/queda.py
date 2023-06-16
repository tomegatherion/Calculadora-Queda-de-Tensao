"""
Formulas

P = (V**2) / Z
P = Z * (I**2)
P = V * I

I = (P / Z) ** 0.5
I = P / V
I = V / Z

Z = P / (I ** 2)
Z (V ** 2) / P
Z = V / I

V = Z * I
V = P / I
V = (P * Z) ** 0.5
"""

resistividade = 0.0178571428571429

ligação = '1FNT'
tensão_entrada = 220
tensão_montante = 210
pot_w = 2200
fp = 1.00
comprimento = 50
bitola = 6

pot_va = pot_w / fp
impedância = (tensão_entrada ** 2) / (pot_w / fp)


if ligação == "1FNT" or ligação =="2FNT":
    fator_k = 200
    corrente = tensão_montante / impedância

else:
    fator_k = 173.2
    corrente = (tensão_montante * 1.73205080757) / impedância

queda_100 = ((fator_k * resistividade * comprimento * corrente) / ((bitola) * tensão_montante))
queda_v = (tensão_montante * (queda_100 / 100))
tensão_jusante = tensão_montante - queda_v

print('Resultados:')
print(f'corrente: {corrente}')
print(f'impedância: {impedância}')
print(f'pot_va: {pot_va}')
print(f'queda %: {queda_100}')
print(f'queda V: {queda_v}')
print(f'tensão a jusante: {tensão_jusante}')