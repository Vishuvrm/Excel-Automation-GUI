class Matrix(list):
    """
    Matrix class is the child of list class.
    It has some additional features like addition, multiplication and Transpose of Matrix objects,
    but features of list are still availabe with the Matrix objects.
    """

    def __init__(self, matr=None):
        # self.__result = [[0 for c in range(len(m1[0]))] for r in range(len(m1))]
        self.__mat = matr
        try:
            self.__trans = [[0 for x in range(self.shape[0])] for y in range(self.shape[1])]
        except:
            self.__trans = [[None], [None]]

    def __add__(self, matr):
        for r in range(len(matr.__mat)):
            for c in range(len(matr.__mat[0])):
                self.__result[r][c] = self.__mat[r][c] + matr.__mat[r][c]

        return self.__result

    def __mul__(self, matr):
        for r in range(len(matr.__mat)):
            for c in range(len(matr.__mat[0])):
                self.__result[r][c] = self.__mat[r][c] * matr.__mat[r][c]

        return self.__result

    @property
    def transpose(self):
        for r in range(len(self.__mat)):
            for c in range(len(self.__mat[0])):
                try:
                    self.__trans[c][r] = self.__mat[r][c]
                except:
                    continue
        self.__mat = self.__trans
        return self.__mat

    @property
    def shape(self):
        try:
            return (len(self.__mat), len(self.__mat[0]))
        except:
            print("result not found!")
    def __str__(self):
        return f"{self.__class__.__name__}({self.__mat})"

    def __repr__(self):
        return f"{self.__class__.__name__}({self.__mat})"