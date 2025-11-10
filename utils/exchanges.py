
class RequiredFieldsError(BaseException):
    def __init__(self, message: str, fields: list[int]):
        super().__init__(message)
        self.fields = fields
        
        
class RowError(BaseException):
    def __init__(self, error: BaseException, row_number: int):
        super().__init__(str(error))
        self.row_number = row_number