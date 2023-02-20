class ColumnExcel:

    def __init__(self, cap, value):
        self.value = 0
        self.cap = cap
        self.before = None
        self.increase_by_value(value)

    def increase(self):
        self.value += 1
        if self.value > self.cap:
            self.value = 1
            if self.before is None:
                self.before = ColumnExcel(self.cap, 1)
            else:
                self.before.increase()

    def increase_by_value(self, amount):
        for i in range(0, amount):
            self.increase()

    def __str__(self):
        column = []
        current = self
        while current is not None:
            column.append(current.value)
            current = current.before
        column.reverse()
        col = ''
        for i in column:
            temp = chr(int(i) + 64)
            col = ''.join([col, temp])
        return col
