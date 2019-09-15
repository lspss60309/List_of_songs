var list = [{
	'id': 377579,
	'name': "浪子回頭",
	'author': "茄子蛋",
	'words_count': 4
},{
	'id': 125535,
	'name': "背叛",
	'author': "蕭靜騰",
	'words_count': 2
}];

function print(){
	var table=document.getElementById("results");	
	var song_keyword = document.getElementById("song_keyword").value;		//欄位中給的歌名關鍵字
	var author_keyword = document.getElementById("author_keyword").value;	//欄位中給的作者關鍵字
	var word_counter = document.getElementById("word_counter").value;		//欄位中給的歌名字數
	
	var color_counter = 0;
	list.forEach(function(i){
		if((word_counter == i.words_count || word_counter == 0 )&& (i.author.indexOf(author_keyword) != -1)&& (i.name.indexOf(song_keyword) != -1)){
			var row = table.insertRow(-1);
			var cell1 = row.insertCell(0);
			var cell2 = row.insertCell(1);
			var cell3 = row.insertCell(2);
			if(color_counter%2 == 0){
				cell1.innerHTML="<div class=\"bg_first\">"+i.id+"</div>";
				cell2.innerHTML=i.name;        
				cell3.innerHTML=i.author;
			}else{
				cell1.innerHTML="<div class=\"bg_second\">"+i.id+"</div>";
				cell2.innerHTML=i.name;        
				cell3.innerHTML=i.author;
			}color_counter++;
		}
	});	
}

function home(){
	window.location.replace('index.html');
}

$(document).ready(function(){
	$("#GoTop").click(function(){
		jQuery("html,body").animate({
			scrollTop:0
		},300);
	});
	$(window).scroll(function() {
		if ( $(this).scrollTop() > 100){
			$('#GoTop').fadeIn("fast");
		} else {
			$('#GoTop').stop().fadeOut("fast");
		}
	});
});
