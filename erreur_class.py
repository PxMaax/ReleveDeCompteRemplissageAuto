class ErreurExcel(Exception):
    def __init__(self, current_cellCoord, details_error):
        super().__init__("Erreur : " + details_error + "case : " + current_cellCoord)
        self.current_cellCoord = current_cellCoord
        self.details_error = details_error
        
        