# Espiral-musical-Excel
A "Espiral musical" em Excel VBA


A “Espiral musical”, a figura abaixo, é construída somente com retas e uma regra simples de ângulos.

![](https://ideiasesquecidas.files.wordpress.com/2021/11/image007.png)

Ela é baseada num vídeo enviado pelo amigo Maurício Cota.

Comece com uma reta qualquer. Depois, trace uma nova reta, adicionando uma rotação com um ângulo.
![](https://ideiasesquecidas.files.wordpress.com/2021/11/reta01.png)


Continue a sequência, agora adicionando reta com 2*ângulo, depois 3*ângulo…
![](https://ideiasesquecidas.files.wordpress.com/2021/11/reta02.png)

Na sexta iteração:
![](https://ideiasesquecidas.files.wordpress.com/2021/11/reta03.png)

Com 100 iterações:
![](https://ideiasesquecidas.files.wordpress.com/2021/11/reta04.png)

O arquivo Excel aqui (https://1drv.ms/x/s!Aumr1P3FaK7jn2acOg89jOQF3Kl4) implementa a rotina, podendo variar ângulos, tamanho da reta e número de iterações.

Dica: Para plotar uma reta no VBA, basta usar o comando abaixo.

    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, x1, y1, x2, y2).Select

Este vai plotar uma reta começando nas coordenadas (x1,y1) e terminando em (x2,y2).

Alguns exemplos:

![](https://ideiasesquecidas.files.wordpress.com/2021/11/espiralmusical02.png)


Mudando um pouco a rotina, é possível fazer degradê de cores.

![](https://ideiasesquecidas.files.wordpress.com/2021/11/espiralmusical03.png)

Outra dica é colocar um ângulo fracionário. Isso porque um ângulo inteiro uma hora vai se tornar periódico, e a figura não será tão legal.

![](https://ideiasesquecidas.files.wordpress.com/2021/11/espiralmusical04.png)

Veja também:

https://ideiasesquecidas.com/2021/11/27/a-espiral-musical-em-excel/


https://ferramentasexcelvba.wordpress.com
