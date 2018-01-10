# Visual Basic .NET

## Contenido

### 1. Conociendo el Lenguaje
#### 1.1 Declaración de variables

```vb.net
Dim texto1 As String
texto1 = "hola mundo"

Dim texto2 As String = "hola mundo!"

Dim num1, num2, num3 As Integer
```

#### 1.2 Tipos de datos

Tipo de Dato | Tamaño | Rango | Declaración
-------------|--------|-------|------------
Short | 16 bits | -32,768 a 32,767 | `Dim num as Short = -3200`
UShort | 16 bits | 0 a 65,535 | `Dim num as UShort = 2500`
Integer | 32 bits | -2,147,483,648 a 2,147,483,647 | `Dim num as UShort = 2500`
UInteger | 32 bits | 0 a 4,294,967,295 | `Dim Seconds As UInteger` <br> `Seconds = 3000000`
Long | 64 bits | -9,223,372,036,854,775,808 a<br>9,223,372,036,854,775,807 | `Dim Bugs As Long` <br>`Bugs = 7800000016`
ULong | 64 bits | 0 a 18,446,744,073,709,551,615 | `Dim SandGrains As ULong` <br>`SandGrains = 1800000000000000000`
Single | 32 bits | -3.4028235E38 a 3.4028235E38 | `Dim UnitCost As Single` <br>`UnitCost = 899.99`
Double | 64 bits | -1.79769313486231E308 a <br>1.79769313486231E308 | `Dim Pi As Double` <br>`Pi = 3.1415926535`
Decimal | 128 bits | 0 a +/-7.923E+28 sin decimales <br>0 a +/- 7.9228162514264337593543950335 (28 decimales) | `Dim Debt As Decimal` <br>`Debt = 7600300.5D` <br>Agregar "D" al inicializar.
Byte | 8 bits | 0 a 255 | `Dim RetKey As Byte RetKey = 13`
SByte | 8 bits | -128 a 127 | `Dim NegNum As SByte NegNum = -20`
Char | 16 bits | Cualquier caracter Unicode en el rango 0–65,535 | `Dim UnicodeChar As Char UnicodeChar = "Ä"c` <br>Agregar "c" al inicializar.
String | Usualmente 16 bits por caracter | 0 a 2,000,000,000 caracteres Unicode | `Dim Greeting As String Greeting = "hello world"`
Boolean | 16 bits | `True` o `False` | Al convertir desde números 0 es `False` y otros valores es `True`.
Date | 64 bits | Enero 1, 0001 a <br>Diciembre 31, 9999 | `Dim Birthday as Date Birthday = #3/17/1900#`
Object | 32 bits | Cualquier tipo puede ser almacenado como tipo Object.<sup>1</sup>| `Dim MyControl As Object`<br>`MyControl = TextBox1`

<sup>1</sup> Adicionalmente, variables de tipo Object pueden contener objetos ya definidos en el proyecto como por ejemplo un TextBox.


#### 1.2 Estructuras de control

##### 1.2.1 If, Else

```vb.net
If var1 = var2 Then
  ...
Else
  ...
End If
```

##### 1.2.2 Select Case

```vb.net
Select Case var1
  Case 1 To 5
    ...
  Case 6
    ...
  Case Else
    ...
End Select
```

##### 1.2.2 Do .. While

```vb.net
While var1 = True
  ...
  ...
Wend
```

##### 1.2.3 Do .. Until

```vb.net
Do
  ...
  ...
Until var1 = True
```

##### 1.2.4 For .. To

```vb.net
For i As Integer = 0 To 10
  ...
Next
```

##### 1.2.5 For Each .. In

```vb.net
For Each obj As Object In lista
  ...
Next
```

#### 1.4 Funciones y procedimientos

##### 1.4.1 Procedimiento

```vb.net
Sub Main()
  ...
End Sub
```

##### 1.4.2 Funciones

```vb.net
Function Calcular(parametro1 As Integer, parametro2 As Interger) As Integer
  Dim resultado as Integer = 0
  ...
  Return resultado
End Funtion
```

#### 1.5 Namespaces

```vb.net
Namespace Padre 
  Module Hijo
  
  End Modulo
End Namespace
```

#### 1.6 Manejo de excepciones

```vb.net
Try
  ...
  Throw New Exception("Error!")
  ...
Catch Ex As Exception
  ...
Finally
  ...
End Try
```

#### 1.7 Arreglos y colecciones

```vb.net
Dim lista(3) As String ' 3 es el máximo índice del arreglo
lista(0) = "Primero"
lista(1) = "Segundo"
lista(2) = "Tercero"
lista(3) = "Cuarto"

Dim numeros() As Integer = {11, 21, 35}
```

#### 1.8 Fechas y otros tipos de datos
### 2. Programación Orientada a Objetos
#### 2.1 Clases
#### 2.2 Propieades
#### 2.4 Métodos
#### 2.5 Constructores
#### 2.6 Herencia
#### 2.7 Interfaces
#### 2.7 
### 3. Windows Forms
#### 3.1 Controles
#### 3.2 Manejo de eventos
#### 3.4 Formularios
### 4. Programación con Bases de Datos

## Ejemplos

## Bibliografía
- **Microsoft Visual Basic 2013 Step by Step** por *Michael Halvorson* - Microsoft Press, 2013
- **Murach's Visual Basic 2012, 5th Edition** por *Anne Boehm* - Mike Murach & Associates, 2010
- **Professional Visual Basic 2012 and .NET 4.5 Programming** por *Bill Sheldon, Billy Hollis, Rob Windsor, David McCarter, Gastón C. Hillar & Todd Herman* - Wiley, 2013
