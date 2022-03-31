 <body bgcolor="#CCCC99" text="#663300" link="#FFFFFF" vlink="#669966" alink="#996666">
<script>

function a_plus_b() {
var a = document.form1.a.value;
var b = document.form1.b.value;
var Hasil = document.form1.Hasil.value;

Hasil== a+b;

}




// End -->
</script> 
<form action="counter.asp" method="post" name="form1" onSubmit="return a_plus_b()">
  <table width="500" border="1">
    <tr bgcolor="#66CCFF"> 
      <td> <div align="right">Bilangan 1</div></td>
      <td> <input name="a" type="text" id="a">
      </td>
    </tr>
    <tr bgcolor="#66CCFF"> 
      <td> <div align="right">Bilangan 2</div></td>
      <td> <input name="b" type="text" id="b"></td>
    </tr>
    <tr bgcolor="#66CCFF"> 
      <td> <div align="right">Hasil </div></td>
      <td> <input name="Hasil" type="text" id="Hasil"  ></td>
    </tr>
    <tr bgcolor="#66CCFF"> 
      <td> <div align="center"> 
          <input type="button" name="Button" value="Get " >
        </div></td>
      <td> <div align="center"></div></td>
    </tr>
  </table>
</form>
