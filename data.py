data = [('CUSTEIO', 171), ('PROTOTIPO, , ', 121), ('PROTOTIPO', 111), ('IMAGEM DE PRODUTO', 104), ('IMAGEM DE PRODUTO, , ', 101), ('CUSTEIO, , ', 81), ('PROTOTIPO, CUSTEIO, ', 38), ('BOOK', 33), ('3D-AMBIENTE, , ', 27), ('REVISÃO DE PREÇO', 25), ('MULTI-MEDIA', 25), ('PRODUÇÃO DE AMOSTRA, , ', 23), ('CERTIFICAÇAO, , ', 20), ('CERTIFICAÇAO', 19), ('REVISÃO DE PREÇO, , ', 18), ('MULTI-MEDIA, , ', 14), ('3D-AMBIENTE', 14), ('BOOK, , ', 13), ('CERTIFICAÇAO, IMAGEM DE PRODUTO, ', 10), ('IMAGEM DE PRODUTO, 3D-AMBIENTE, ', 10), ('VIDEO, , ', 9), ('IMAGEM DE PRODUTO, CUSTEIO, ', 8), ('PRODUÇÃO DE AMOSTRA', 8), (', , ', 7), ('CUSTEIO, PROTOTIPO, ', 6), ('CUSTEIO, IMAGEM DE PRODUTO, ', 5), ('VIDEO', 5), ('CERTIFICAÇAO, IMAGEM DE PRODUTO, CUSTEIO', 4), ('IMAGEM DE PRODUTO, CERTIFICAÇAO, REVISÃO DE PREÇO', 3), ('PRODUÇÃO DE AMOSTRA, CUSTEIO, ', 3), ('DESENV. PROD.', 3), ('REENGENHARIA', 3), ('DESENV. PROD., , ', 2), ('CERTIFICAÇAO, PROTOTIPO, ', 2), ('EXPLORAÇÃO RESTYLING | Homólogo , CUSTEIO, IMAGEM DE PRODUTO', 2), ('CUSTEIO, PRODUÇÃO DE AMOSTRA, ', 2), ('DESENV. PROD., PROTOTIPO, ', 1), ('BOOK, IMAGEM DE PRODUTO, CERTIFICAÇAO', 1), ('CERTIFICAÇAO, IMAGEM DE PRODUTO, 3D-AMBIENTE', 1), ('EXPLORAÇÃO RESTYLING | Homólogo , , ', 1), ('IMAGEM DE PRODUTO, VIDEO, ', 1), ('IMAGEM DE PRODUTO, VIDEO, PROTOTIPO', 1), ('PROTOTIPO, CERTIFICAÇAO, ', 1), ('CUSTEIO, DESENV. PROD., ', 1), ('IMAGEM DE PRODUTO, CUSTEIO, PROTOTIPO', 1), ('DESENV. PROD., CERTIFICAÇAO, ', 1), ('PROTOTIPO, IMAGEM DE PRODUTO, ', 1), ('3D- MODELAÇÃO, , ', 1), ('PROTOTIPO + CERTIFICAÇÃO', 1), ('EXPLORAÇÃO RESTYLING | Homólogo ', 1)]



parse_data = {}
clean_request = []

for d in data:
    request = d[0].strip().split(',')
    value = d[1]
    for x in request:
        cleaner = x.strip()
        if len(cleaner) > 0:
            clean_request.append([cleaner.upper(), value])
    


for j in clean_request:
    request_clean = j[0]
    value_clean = int(j[1])
    if request_clean not in parse_data:
        parse_data[request_clean] = value_clean
    else:
        parse_data[request_clean] += value_clean
    
total = 0
for values in parse_data:
    print(f"Request Type: {values.capitalize()} - {parse_data[values]} DRS´s")
    total += parse_data[values]
print(f'Total: {total}')
        




