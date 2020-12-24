import os

print("Hoşgeldiniz")
print("Bir seçenek seçiniz :")
print("1.Fatura yazdırma \n"
    "2.E-mail gönderme")
choice = input("Bir rakam girip enter a basınız :")

if(choice == "1" or choice == " 1"):
    print("Seçenek 1 seçildi")
elif(choice == "2" or choice == " 2"):
    print("Seçenek 2 seçildi")
else:
    print("Lütfen geçerli bir seçenek giriniz (Örn: '1' veya '2'")


