# Module Player

class Player:
    currentid=0
    
    def __init__(self, category) -> None:
        self.id=Player.currentid
        currentid+=1
        self.category=category
        pass

    def __eq__(self, __o: object) -> bool:
        return self.id==__o.id 
        pass

