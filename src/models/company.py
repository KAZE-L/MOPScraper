class Company:
    """公司資料模型類別"""
    def __init__(self, name=None, code=None):
        """初始化公司資料"""
        self.name = name
        self.code = code
        self.industry = None
        self.annual_revenue = None
        self.gross_profit = None
        self.gross_margin = None
        self.profit_before_tax = None
        self.profit_after_tax = None

    def to_dict(self):
        """轉換為字典格式"""
        return {
            '公司名稱': self.name,
            '公司代號': self.code,
            '產業類別': self.industry,
            '年營收': self.annual_revenue,
            '毛利額': self.gross_profit,
            '毛利率': self.gross_margin,
            '稅前淨利': self.profit_before_tax,
            '稅後淨利': self.profit_after_tax
        }

    @classmethod
    def from_dict(cls, data):
        """從字典創建實例"""
        company = cls(data.get('公司名稱'), data.get('公司代號'))
        company.industry = data.get('產業類別')
        company.annual_revenue = data.get('年營收')
        company.gross_profit = data.get('毛利額')
        company.gross_margin = data.get('毛利率')
        company.profit_before_tax = data.get('稅前淨利')
        company.profit_after_tax = data.get('稅後淨利')
        return company 