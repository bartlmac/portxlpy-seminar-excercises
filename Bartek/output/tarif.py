# Automatisch generiert
def raten_zuschlag(zw):
    return IF(zw=2,2%,IF(zw=4,3%,IF(zw=12,5%,0)))
