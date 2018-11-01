M2000 Interpreter and Environment

Version 9.4 rev 24 active-X

Works with infinity positive and negative. To make a constant use this:

      Function Infinity(positive=True) {
            buffer clear inf as byte*8
            m=0x7F
            if not positive then m+=128
            return inf, 7:=m, 6:=0xF0
            =eval(inf, 0 as double)
      }
      K=Infinity(false)
      L=Infinity()
      Function TestNegativeInfinity(k) {
            =str$(k, 1033) = "-1.#INF"
      }
      Function TestPositiveInfinity(k) {
            =str$(k, 1033) = "1.#INF"
      }
      Function TestReturn$ {
            =str$(Number, 1033)
      }
 
      Print TestNegativeInfinity(K), TestPositiveInfinity(L)


From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate

https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

http://georgekarras.blogspot.gr/

http://m2000.forumgreek.com/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             