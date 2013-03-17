#levelname(BattleStation Devastation)
#levelnum(4)
#levelduration(158500)
#leveldescbkpic(mars.bmp)
#leveldescduration(30000)
#levelbrief {
 Earth forces have delpoyed a large Orbital BattleStation to counter enemy attacks, but
 it needs time to activate it.Enemies are closing up...defenend the BattleStation until
 it's ready!
}


#levelstate(1,10500) {
#addsmq(6000,,,,,,Defend the BattleStation at any cost.)
#addsmq(4500,,,,,,It's our last hope!)
#createbs(0)
}


#levelstate(2,40000) {
#timewarpship(8500,9500,random,3,1)
#timewarpship(2500,3000,random,2,1)
#timewarpship(4200,4900,random,1,1)
}

#levelstate(3,8000) {
#addsmq(6000,,,,,,Meteors approaching!)
}

#levelstate(4,20000) {
#timewarpmeteor(4050,4100,2,random)
#timewarpship(2700,3000,random,1,1)
}

#levelstate(5,9000) {
}

#levelstate(6,10000) {
#createbs(1)
#addsmq(4000,,,,,,BattleStation is ready.)
#addsmq(6000,,,,,,Sensors detect large ship mass.)
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
#addsmq(10000,,,,,,Thanks to you once again the Earth is safe.)
}

