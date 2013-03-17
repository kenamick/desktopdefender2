#levelname('Battlestation Devastation')
#levelnum(4)
#levelduration(158500)
#leveldescbkpic(mars.bmp)
#leveldescduration(60000)
#levelbrief {
 Our scientists have managed to deploy a BattleStation into Earth's
orbit, but they will need some time to activate it.
In the meantime our sensors detected a new, larger fleet of the
Black League.
Do not let the enemies destroy the BattleStation.It could be our 
only salvation...
--------------------------------------------------------------------
Mission: Defend the BattleStation until the scientists get it ready!
}


#levelstate(1,10500) {
#addsmq(6000,,,,,,Defend the BS at all costs.)
#addsmq(4500,,,,,,It's our last hope!)
#createbs(0)
}


#levelstate(2,40000) {
#timewarpship(9500,10500,random,3,1)
#timewarpship(3000,3200,random,2,1)
#timewarpship(5000,5200,random,1,1)
}

#levelstate(3,8000) {
#addsmq(6000,,,,,,Meteors are closing in!)
}

#levelstate(4,20000) {
#timewarpmeteor(4050,4100,2,random)
#timewarpship(3200,3250,random,1,1)
}

#levelstate(5,9000) {
}

#levelstate(6,10000) {
#createbs(1)
#addsmq(4000,,,,,,The BattleStation is ready!)
#addsmq(6000,,,,,,Sensors detect large enemy fleet.)
}

#levelstate(7,40000) {
#timewarpship(20000,22000,random,4,1)
#timewarpship(7000,8000,random,1,2)
#timewarpship(5600,7600,random,2,1)
#timewarpship(14000,16000,random,3,1)
}

#levelstate(8,8000) {
}

#levelstate(9,12000) {
#addsmq(10000,,,,,,Once again the Earth is saved.)
}

