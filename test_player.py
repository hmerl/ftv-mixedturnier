import player

def test_playercreation():

    oneplayer=player.Player(1)
    assert oneplayer.id>-1
    assert oneplayer.category==1
    assert oneplayer.currentid>oneplayer.id

