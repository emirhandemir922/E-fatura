import os
import Fatura
import Mail

print("Hoşgeldiniz")
print("Bir seçenek seçiniz :")
print("1.Fatura yazdırma \n"
    "2.E-mail gönderme")
choice = input("Bir rakam girip enter a basınız :")

if(choice == "1" or choice == " 1"):
    os.system('Fatura.py')
elif(choice == "2" or choice == " 2"):
    exec(open("Mail.py").read())
else:
    print("Lütfen geçerli bir seçenek giriniz (Örn: '1' veya '2'")
