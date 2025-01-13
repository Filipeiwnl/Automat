class DesignatorData:
    def __init__(self, tx_a, rx_a, ponta_a, ponta_b, db_km_a, db_km_b):
        self.tx_a = tx_a
        self.rx_a = rx_a
        self.ponta_a = ponta_a
        self.ponta_b = ponta_b
        self.db_km_a = db_km_a
        self.db_km_b = db_km_b

    def to_dict(self):
        return {
            'TX A': self.tx_a,
            'RX A': self.rx_a,
            'Ponta A': self.ponta_a,
            'Ponta B': self.ponta_b,
            'DB/KM A': self.db_km_a,
            'DB/KM B': self.db_km_b
        }
