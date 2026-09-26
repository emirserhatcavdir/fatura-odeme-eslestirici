fatura_kurus = 500000  # 5.000 TL
odeme_kurus = 300000   # 3.000 TL

fark_kurus = fatura_kurus - odeme_kurus

if fark_kurus > 0:
    print("Eksik ödeme:", fark_kurus // 100, "TL")
elif fark_kurus == 0:
    print("Fatura tam ödendi.")
else:
    fazla_kurus = odeme_kurus - fatura_kurus
    print("Fazla ödeme:", fazla_kurus // 100, "TL")
