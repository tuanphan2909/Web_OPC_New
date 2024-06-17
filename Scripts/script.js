




fetch("http://118.69.225.144/api/user")
.then(function(response){
	return response.json();
})
.then(function(products){
	let placeholder = document.querySelector("#data-output");
	let out = "";
	for(let product of products){
		out += `
			<tr>
				
				<td>${product.tendangnhap}</td>
				<td>${product.matkhau}</td>
				<td>${product.ma_Dvcs}</td>
			
			</tr>
		`;
	}

	placeholder.innerHTML = out;
});
const searchFun = () =>{
	let filter = document.getElementById('myInput').value.toUpperCase();
	let myTable = document.getElementById('myTable');
	let tr = myTable.getElementsByTagName('tr');
	  for(var i =0;i<tr.length;i++){
	  let td = tr[i].getElementsByTagName('td')[0];
	  if(td)
	  {
		let textValue = td.textContent || td.innerHTML;
		if(textValue.toLocaleUpperCase().indexOf(filter)>-1){
			tr[i].style.display="";

		}else{
			tr[i].style.display="none";
		}
	  }
  }
  
  }
