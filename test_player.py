import player

def test_playercreation():

    oneplayer=player.Player(1)
    assert oneplayer.id>-1
    assert oneplayer.category==1
    assert oneplayer.currentid>oneplayer.id

def test_playercomparison():
    firstplayer=player.Player(1)
    secondplayer=player.Player(1)

    assert firstplayer==firstplayer
    assert secondplayer==secondplayer
    assert firstplayer!=secondplayer
    assert secondplayer!=firstplayer
