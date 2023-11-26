from math import e

def calculate_forward_price(spot: float, interest_rate: float, time: float):
    return round(spot * e ** (interest_rate * time),2)

