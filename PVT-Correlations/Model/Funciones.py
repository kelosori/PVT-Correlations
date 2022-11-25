
import math

#Funcion de la uo
def uo(colums, API, T):
    """

    Parameters

    colums
        Correlación usada (Beggs y robinson o Glaso)
    API
        Gravedas API
    T
        Temperatura en °R
    Returns
        uo en cp


    """
    z = 3.0324 - (0.02023 * API)
    y = 10**z
    x = (y) * (T-460)**(-1.163)
    a1=(10.313 * math.log10(T-460)) - 36.447
    if colums == "Beggs & Robinson":
        uo= 10**x - 1
    else:
        uo= 3.141 * (10**10) * (T-460)**(-3.414) * (math.log10(API))**a1
    return uo


#Funcion PB
def Pb(colums, Rs, Yg, T, API=None, Yo=None):
    """
    Parameters
    ----------
    colums
        Correlación usada
    Rs
        Solubilidad del gas en PCN/BN
    Yg
        Solubilidad del gas
    T
        Temperatura en °R
    API
        Gravedad API
    Yo
        Gravedad específica del petroleo
    Returns
    Pb en psi
    -------
    """
    a= 0.00091 * (T-460) - 0.0125*(API)
    a1=5.38088 * 10**(-3)
    b=0.715082
    c=-1.87784
    d=3.1437
    e=1.32657
    if colums == "Standing":
        Pb = 18.2 * (((Rs / Yg) ** 0.83) * 10 ** a - 1.4)
    else:
        Pb=(a1 * Rs**b) * (Yg**c) * (Yo**d) * (T**e)
    return  Pb

#Funcion del Bo

def Bo(colums,Rs, Yg, Yo, T ):
    """"
    Parameters
    ---------------
    colums:
        Correlación usada
    Rs:
        Solubilidad del gas en PCN/BN
    Yg:
        Gravedad específica del gas en solución
    Yo:
        Gravedad específica del crudo en superficie
    T:
        Temperatura del sistema en °R
    Returns
        float number  -> Bo
    """
    a=0.7423
    b=0.3232
    c=-1.202
    F = (Rs**a) * (Yg**b) * (Yo**c)

    if colums == "Standing":
        Bo = 0.9759 + 0.00012 * (Rs*((Yg/Yo)**0.5) + 1.25*(T-460))**(1.2)
    else:
        Bo = 0.497069 + 0.862963 * (10**(-3)) * T + 0.182594 * 10**(-2) * F + 0.318099 * 10**(-5) * F**(2)

    return Bo

#Funcion para encontrar el Rs
def Rs(colums, P, API, T=None, Yg=None, Yo=None):
    """"
    Parameters
    ---------------
    colums:
        Correlación usada
    P:
        Presión del sistema en psi
    API:
        Gravedad API
    T:
        Temperatura del sistema en °R
    Yg:
        Gravedad específica del gas en superficie
    Yo:
        Gravedad específica del crudo en superficie
    Returns
        float number  -> Rs
    """
    x = 0.0125 * API - 0.00091 * (T - 460)
    a = 185.843208
    b = 1.877840
    c = -3.1437
    d = -1.32657
    e = 1.398441
    if colums == "Standing":
        Rs = (Yg) * (((P / 18.2) + 1.4) * (10 ** x)) ** 1.2048
    else:
        Rs = (a * (Yg ** b) * (Yo ** c) * (T ** d) * P) ** e
    return Rs