function scroll_club(){
$(function(){
      $('#showArea li').eq(0).fadeOut('slow',function(){        
            $(this).clone().appendTo($(this).parent()).fadeIn('slow');
            $(this).remove();
        });
     });
}
setInterval('scroll_club()',3000);

//$(document).ready(function(){  
	$("#searchTabs a").click(function(){
		$("#searchTabs li").removeClass("current");
		$(this).parent().addClass("current");
		$("#search").attr("value",$(this).html());
		return false;
	})
	
//滚动三条
	var PMarquee = BGet("srollingTips") //滚动对象
	if (PMarquee){	
		var PLineHeight = 20; //单行高度，像素
		var lis = PMarquee.getElementsByTagName("span");
		var PLineCount = lis.length; //实际行数
		var PScrollAmount = 20; //每次滚动高度，像素
		
		function srun() {
			PMarquee.scrollTop ++;
			if (PMarquee.scrollTop == PLineCount * PLineHeight ) {PMarquee.scrollTop = 0;}
			if (PMarquee.scrollTop % PLineHeight == 0 ) {window.setTimeout( "srun()", 3000 ); }//滚动间歇控制
			else {window.setTimeout( "srun()", 20 ); }//滚动速度控制
		}
		if (PLineCount>1){
		PMarquee.innerHTML += PMarquee.innerHTML;
		window.setTimeout( "srun()", 5000 ); //滚动间歇控制
		}
	}	

//幻灯片
if (BGet("slide")){
	var theInt = null;
	var curclicked = 0;
	$(function(){
		$('#pic_list a').click(function(){return false});
		t(0);
		$('#pic_list span').click(function(){
			if($('#this_pic').attr('src') == $(this).attr('lang')) return;
			t($('#pic_list span').index($(this)));
			$('#pic_list span').removeClass("current")
			$(this).addClass("current")
		});
	});
	t = function(i){
		clearInterval(theInt);
		if( typeof i != 'undefined' )
		curclicked = i;
			$('#this_pic').hide().attr('src',$('#pic_list span').eq(i).attr('lang')).fadeIn(750);
			$('#this_a').attr('href',$('#pic_list span').eq(i).parents('a').attr('href'));
			$('#pic_list span').removeClass("current")
			$('#pic_list span').eq(i).addClass("current")
		theInt = setInterval(function (){
			i++;
			if (i > $('#pic_list span').length - 1) {i = 0};
			$('#this_pic').hide().attr('src',$('#pic_list span').eq(i).attr('lang')).fadeIn(750);
			$('#this_a').attr('href',$('#pic_list span').eq(i).parents('a').attr('href'));
			$('#pic_list span').removeClass("current")
			$('#pic_list span').eq(i).addClass("current")
		},4000)
	}	
}	
//滚动图册
	transpics()

//登录区点击切换		
	$("#loginTips p").hide();
	$("#loginTips p").eq(0).show();
	$("#titles span").click(function(){
		$("#titles span").removeClass("current");
		$(this).addClass("current");
		$("#loginTips p").hide();
		$("#loginTips p").eq($("#titles span").index($(this))).show();
		
	})
	
//活页夹切换	
	BtTabOn("ca_head","ca")
	BtTabOn("area_head","area")

//排名前三的自动变亮	
	$("#rank01 li em").eq(0).addClass("front")
	$("#rank01 li em").eq(1).addClass("front")
	$("#rank01 li em").eq(2).addClass("front")
	$("#rank02 li em").eq(0).addClass("front")
	$("#rank02 li em").eq(1).addClass("front")
	$("#rank02 li em").eq(2).addClass("front")
	$("#rank03 td em").eq(0).addClass("front")
	$("#rank03 td em").eq(1).addClass("front")
	$("#rank03 td em").eq(2).addClass("front")
//})	

//操作排序
	$("#orders span").mouseover(function(){$(this).addClass("on")}).mouseout(function(){$(this).removeClass("on")})
//显示二级菜单
<!--//--><![CDATA[//><!--
function menuFix() {
 var sfEls = document.getElementById("nav").getElementsByTagName("li");
 for (var i=0; i<sfEls.length; i++) {
  sfEls[i].onmouseover=function() {
  this.className+=(this.className.length>0? " ": "") + "sfhover";
  }
  sfEls[i].onMouseDown=function() {
  this.className+=(this.className.length>0? " ": "") + "sfhover";
  }
  sfEls[i].onMouseUp=function() {
  this.className+=(this.className.length>0? " ": "") + "sfhover";
  }
  sfEls[i].onmouseout=function() {
  this.className=this.className.replace(new RegExp("( ?|^)sfhover\\b"), 

"");
  }
 }
}
window.onload=menuFix;

//--><!]]>

//推荐间隔背景色
	$("#zh_recommed .zh_r_list:odd").css("background","#F8F8F8")