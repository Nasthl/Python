import requests

url = 'https://zerotojunior.dev/cezar.txt'

response = requests.get(url)

# Sprawdź, czy żądanie zakończyło się sukcesem (kod odpowiedzi 200)
if response.status_code == 200:
    # Odczytaj zawartość pliku, uwzględniając odpowiednie kodowanie znaków
    text = response.content.decode('UTF-8', errors='ignore')
    print("Pobrany tekst:")
    print(text)
else:
    print(f"Błąd pobierania pliku: {response.status_code}")


small = "aąbcćdeęfghijklłmnńoóprsśtuwyzźż"
large = "AĄBCĆDEĘFGHIJKLŁMNŃOÓPRSŚTUWYZŹŻ"
for j in range(len(small)):
    wynik = ''
    for i in text:
        if small.find(i)>=0:
            wynik=wynik+small[(small.find(i)+j)%len(small)]
        elif large.find(i)>=0:
            wynik=wynik+large[(large.find(i)+j)%len(large)]
        else:
            wynik=wynik+i
    test=[]
    wynik_test=wynik.lower()
    for i in small:
        test.append(wynik_test.count(i))
    if test.index(max(test))==6:
        break
print(wynik)
