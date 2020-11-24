import enum


class Section(enum.Enum):
    Summary = 0
    LegalTitle = 1
    Lease = 2
    SearchResult = 3
    AdditionalInformation = 4
    Survey = 5
    StampLandTax = 6
    Mortgage = 7
    Documents = 8
    Exchange = 9
    Financial = 10
    Conclusion = 11

    def __lt__(self, other):
        return self.value < other.value
