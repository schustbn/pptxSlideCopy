class Relationship:
    """Relationship."""

    def __init__(self, id: str, type: str, target: str, targetMode=""):
        """Init."""
        self.id = id
        self.type = type
        self.target = target
        self.targetMode = targetMode

    def addName(self, name):
        """Add name."""
        # print ('Name: ' + name)
        self.name = name
