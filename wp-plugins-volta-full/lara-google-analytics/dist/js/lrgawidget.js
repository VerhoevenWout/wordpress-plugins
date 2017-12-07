
/**
 * @package    Google Analytics by Lara
 * @author     Amr M. Ibrahim <mailamr@gmail.com>
 * @link       https://www.whmcsadmintheme.com
 * @copyright  Copyright (c) WHMCSAdminTheme 2016
 */

window.gauthWindow = function (url) {
      var newWindow = window.open(url, 'name', 'height=600,width=450');
      if (window.focus) {
        newWindow.focus();
      }
}

window.debugWindow = function () {
      var newWindow = window.open('', 'Debug', 'height=600,width=600,scrollbars=yes');
	  newWindow.document.write("<pre>"+JSON.stringify(lrgawidget_debug, null, " ")+"</pre>");
      if (window.focus) {
        newWindow.focus();
      }
}

window.lrgawidget_debug;

(function($) {

	
var dateRange = {};
var lrsessionStorageReady = false;
var setup = false;
var debug = false;


function isObject(val) {
    if (val === null) { return false;}
    return ( (typeof val === 'function') || (typeof val === 'object') );
}

function reloadCurrentTab(){
   var $link = $('#lrgawidget li.active a[data-toggle="tab"]');
   $link.parent().removeClass('active');
   var tabLink = $link.attr('href');
   $('#lrgawidget a[href="' + tabLink + '"]').tab('show');	
}

function lrgaErrorHandler(err){
	var error;
	var error_description;
	var error_code;
	var error_debug;
	var message;
	if (typeof err === 'object'){
		error = ((err.error != null) ? "["+err.error+"]" : "");
		error_description = ((err.error_description != null) ? err.error_description : "");
		error_code = ((err.code != null) ? "code ["+err.code+"]" : "");	
		if (err.debug != null){
			error_debug = "<a href='javascript:debugWindow();'>debug</a>";
			lrgawidget_debug = err.debug;
		}
        message = "Error "+error_code+" "+error_debug+":<br> "+error+" "+error_description;
	}else {
		message = err;
	}
    $("#lrgawidget_error").html('<h4><i class="icon fa fa-ban"></i> '+message+'</h4>');
	$("#lrgawidget_error").removeClass("hidden");	
}

function lrWidgetSettings(arr){
	$("#lrgawidget_error").html("").addClass("hidden");
	$("#lrgawidget_mode").html("");
	$("#lrgawidget_loading").html('<i class="fa fa-spinner fa-pulse"></i>');

	if (arr[0]){
		arr[0].value = "lrgawidget_"+arr[0].value;
	}else{
		arr['action'] = "lrgawidget_"+arr['action'];
	}
	if (typeof arr === 'object'){
		try {
			arr.push({name: 'start', value: dateRange.start});
			arr.push({name: 'end', value: dateRange.end});
		}catch(e){
			arr['start'] = dateRange.start;
			arr['end'] = dateRange.end;
		}
	}

	if (debug){console.log(arr)};
	return $.ajax({
		method: "POST",
		url: lrgawidget_ajax_object.lrgawidget_ajax_url,
		data: arr,
		dataType: 'json'
	})
	.done(function (data, textStatus, jqXHR) {
		if (debug){console.log(data)};
		if (data.status != "done"){
			lrgaErrorHandler(data);
		}
		
		if (data.setup === 1){
			setup = true;
			if ($("[href='#lrgawidget_settings_tab']").is(":visible")){
				$("[href='#lrgawidget_settings_tab']").tab('show');
			}else{
				lrgaErrorHandler("Initial Setup Required!<br><br>Please contact an administratior to complete the widget setup.");
			}
		}
		
		if (data.status == "done"){
			if (data.cached){ $("#lrgawidget_mode").html('cached');}
		}		
	})
	.fail(function (jqXHR, textStatus, errorThrown) {
		lrgaErrorHandler(errorThrown);
		if (debug){
			console.log(jqXHR);
			console.log(textStatus);
			console.log(errorThrown);
		}
	})		
	.always(function (dataOrjqXHR, textStatus, jqXHRorErrorThrown) {
		$("#lrgawidget_loading").html("");
		$("#lrgawidget_loading_big").hide();
	});
}


var lrgaProfiles;
var lrgaaccountID; 
var webPropertyID;
var profileID;
var webPropertyUrl;



function populateViews(){
	$('#lrgawidget-accounts').html("");
	$('#lrgawidget-properties').html("");
	$('#lrgawidget-profiles').html("");
    
	$.each(lrgaProfiles, function( index, account ) {
		if (account.id){
			if (!lrgaaccountID){lrgaaccountID = account.id;}
			$('#lrgawidget-accounts').append($("<option></option>").attr("value",account.id).text(account.name)); 
			if (account.id == lrgaaccountID){
				$("#lrgawidget-accname").html(account.name);
				if (account.webProperties){
					$.each(account.webProperties, function( index, webProperty ) {
						if (!webPropertyID){webPropertyID = webProperty.id;}
						$('#lrgawidget-properties').append($("<option></option>").attr("value",webProperty.id).text(webProperty.name + " - [ " + webProperty.id + " ] "));
						if (webProperty.id == webPropertyID){
							$("#lrgawidget-propname").html(webProperty.name);
							$("#lrgawidget-propurl").html(webProperty.websiteUrl+ " - [ " + webProperty.id + " ] ");
							webPropertyUrl = webProperty.websiteUrl;
							if (webProperty.profiles){
								$.each(webProperty.profiles, function( index, profile ) {
									if (!profileID){profileID = profile.id;}
									$('#lrgawidget-profiles').append($("<option></option>").attr("value",profile.id).text(profile.name));
									if (profile.id == profileID){
										$("#lrgawidget-vname").html(profile.name);
										$("#lrgawidget-vtype").html(profile.type);
										
									}
								});
							}
						}											 
					});
				}
			}
		}
	});


	$('#lrgawidget-accounts').val(lrgaaccountID);
	$('#lrgawidget-properties').val(webPropertyID);
	$('#lrgawidget-profiles').val(profileID);
	
}

$(document).ready(function(){
	
    $("#lrgawidget-credentials").submit(function(e) {
        e.preventDefault();
		lrWidgetSettings($("#lrgawidget-credentials").serializeArray()).done(function (data, textStatus, jqXHR) {
			$('#lrga-wizard').wizard('selectedItem', {step: "lrga-getCode"});
			$('#lrga-wizard #code-btn').attr('href','javascript:gauthWindow("'+data.url+'");');
			$('#lrgawidget-code input[name="client_id"]').val($('#lrgawidget-credentials input[name="client_id"]').val());
			$('#lrgawidget-code input[name="client_secret"]').val($('#lrgawidget-credentials input[name="client_secret"]').val());
		})
	});
	
	
    $("#lrgawidget-code").submit(function(e) {
        e.preventDefault();
		lrWidgetSettings($("#lrgawidget-code").serializeArray()).done(function (data, textStatus, jqXHR) {
			if (data.status == "done"){
				$('#lrga-wizard').wizard('selectedItem', {step: "lrga-profile"});
			}
		})
	});	
	
    $("#express-lrgawidget-code").submit(function(e) {
        e.preventDefault();
		lrWidgetSettings($("#express-lrgawidget-code").serializeArray()).done(function (data, textStatus, jqXHR) {
			if (data.status == "done"){
				$('#lrga-wizard').wizard('selectedItem', {step: "lrga-profile"});
			}
		})
	});		
	
	
    $("#lrgawidget-setProfileID").submit(function(e) {
        e.preventDefault();
		lrWidgetSettings($("#lrgawidget-setProfileID").serializeArray()).done(function (data, textStatus, jqXHR) {
			if (data.status == "done"){
				$("#lrgawidget a[href^='#lrgawidget_']:eq(1)").click();
			}
		})	
	});		
	
	$('#lrga-wizard').on('changed.fu.wizard', function (evt, data) {
		if ($("[data-step="+data.step+"]").attr("data-name") == "lrga-profile"){
			lrWidgetSettings({action: "getProfiles"}).done(function (data, textStatus, jqXHR) {
				if (data.status == "done"){
					lrgaaccountID = data.current_selected.account_id;
					webPropertyID = data.current_selected.property_id;
					profileID = data.current_selected.profile_id;
					lrgaProfiles = data.all_profiles;
					populateViews();
					setup = false;
				}
			})
		}
	});
	
	$('#lrgawidget-accounts').on('change', function() {
		lrgaaccountID = this.value;
		webPropertyID = "";
		profileID = "";
		populateViews();
	});

	$('#lrgawidget-properties').on('change', function() {
		webPropertyID = this.value;
		profileID = "";
		populateViews();
	});	

	$('#lrgawidget-profiles').on('change', function() {
		profileID = this.value;
		populateViews();
	});	
	$('a[data-reload="lrgawidget_reload_tab"]').on('click', function(e) {
		 e.preventDefault();
		 reloadCurrentTab();
	});
	
	$('a[data-reload="lrgawidget_go_advanced"]').on('click', function(e) {
		 e.preventDefault();
		 $("#lrgawidget_express_setup").hide();
		 $("#lrgawidget_advanced_setup").show();
		 $("[data-reload='lrgawidget_go_express']").show();
		 
	});	
	
	$('[data-reload="lrgawidget_go_express"]').on('click', function(e) {
		 e.preventDefault();
		 $("#lrgawidget_error").html("").addClass("hidden");
		 $("#lrgawidget_advanced_setup").hide();
		 $("[data-reload='lrgawidget_go_express']").hide();
		 $('#lrga-wizard').wizard('selectedItem', {step: 1});
		 $("#lrgawidget_express_setup").show();
	});	
	
});


var mainChart;
var mainChartDefaults = {"grid"   : {axisMargin: 20, hoverable: true, borderColor: "#f3f3f3",	borderWidth: 1,	tickColor: "#f3f3f3", mouseActiveRadius: 350},
						 "series" : {shadowSize: 1},
						 "lines"  : {fill: true, color: ["#3c8dbc", "#f56954"]},
						 "yaxes"  : [{ min: 0 }],
						 "xaxis"  : {mode: "time",timeformat: "%b %d"},
						 "colors" : ["#3c8dbc"],
						 "legend" : {show: true, container:'#lrga-legendholder'}};
						 
var lastFlotIndex = null;
var currentPlotData = {};

function lrTickFormatter (val, axis){
	if(Math.round(val) !== val) { val = val.toFixed(2);}
	return axis.options.lrcustom.before + val;
}

function lrLegendFormatter(label, series){
   if (series.lrcustom.total >= 0){
	   return label+"<span class='pull-right' style='margin-left: 10px;'><b>"+series.lrcustom.before + " " + series.lrcustom.total+"</b></span>";
   }
}						 

function drawGraph(data,name){
	var settings = mainChartDefaults;
	var gData = [{ data:data["data"], label:data["label"], lines: { show: true },points: { show: true}, lrcustom: {before: data["lrbefore"], after: data["lrafter"], format: data["lrformat"]}}];
	$("#lrgawidget_sessions_chart_tooltip").remove();
	$("#lrgawidget_sessions_chartDiv").removeData("plot").empty();

	if (mainChart){
        mainChart.shutdown();
		mainChart.destroy();
        mainChart = null;
		lastFlotIndex = null;
		currentPlotData = {};
	}

	if (gData.length < 2){$("#lrga-legendholder").css({"right":"17px"})};

	mainChart = $.plot($("#lrgawidget_sessions_chartDiv"), gData, settings);
	currentPlotData = mainChart.getData();	
	
	$('<div class="tooltip-inner" id="lrgawidget_sessions_chart_tooltip"></div>').css({
		"text-align": "left",
		"position": "absolute",
		"display": "none",
		"opacity": 0.8
	}).appendTo("body");

	$("#lrgawidget_sessions_chartDiv").bind("plothover", function (event, pos, item) {
		if (item) {
			if  ((lastFlotIndex != item.dataIndex)){
				lastFlotIndex = item.dataIndex;
				if (debug){ console.log(item);
							console.log(currentPlotData);
							console.log(lastFlotIndex);
					}
				var x = item.datapoint[0].toFixed(2);
				var y = item.datapoint[1];
				var rightMargin = "auto";
				var leftMargin  = "auto";
				var formattedDateString = moment(item.datapoint[0]).format('ddd, MMMM D, YYYY');
				
				var currToolTipText = formattedDateString + "<br>";
				$.each(currentPlotData, function( i, dSeries ){
					var cItem = dSeries.data[item.dataIndex][1];
					if (dSeries.lrcustom.format == "seconds"){ cItem = formatSeconds(cItem);}
					currToolTipText += '<div style="display: inline-block;padding:1px;"><div style="width:4px;height:0;border:4px solid '+dSeries.color+';overflow:hidden"></div></div><div style="display: inline-block;padding-left:5px;">'+dSeries.label+' : '+dSeries.lrcustom.before + cItem + dSeries.lrcustom.after +"</div><br>";
				});
				
				if(item.pageX + 250 > $(document).width()){ 
					rightMargin = ($(document).width() - item.pageX) + 15;
				}else{
					leftMargin  = item.pageX + 15;
				}
				
				$("#lrgawidget_sessions_chart_tooltip").html(currToolTipText)
					.css({top: item.pageY - 25, left: leftMargin, right: rightMargin})
					.show();
			}
		} else {
			lastFlotIndex = null;
			$("#lrgawidget_sessions_chart_tooltip").hide();
			$("#lrgawidget_sessions_chart_tooltip").empty();
		}
	});
}

function formatSeconds(totalSec){
	var hours   = Math.floor(totalSec / 3600);
	var minutes = Math.floor((totalSec - (hours * 3600)) / 60);
	var seconds = totalSec - (hours * 3600) - (minutes * 60);
	var fseconds = seconds.toFixed(0);
	var result = (hours < 10 ? "0" + hours : hours) + ":" + (minutes < 10 ? "0" + minutes : minutes) + ":" + (fseconds  < 10 ? "0" + fseconds : fseconds);	
	return result;
}

function drawSparkline(id, data, color){
	if (!color){color = '#b1d1e4';}
	$(id).sparkline(data.split(','), {
		type: 'line',
		lineColor: "#3c8dbc",
		fillColor: color,
		spotColor: "#3c8dbc",
		minSpotColor: "#3c8dbc",
		maxSpotColor: "#3c8dbc",
		drawNormalOnTop: false,
		disableTooltips: true,
		disableInteraction: true,
		width:"100px"
		});
}

var plotData = {};
var plotTotalData = {};
var selectedPname = "";

function drawMainGraphWidgets(data, selected){
	if ($('#lrgawidget_sb-main').is(":visible")){
		$.each(data, function( name, raw ){
			var appnd = "";
			var color = "";
			if ((name == "percentNewSessions") || (name == "bounceRate")){ appnd = " %";} 
			if (name == selected ){  color = "#77b2d4";} 		
			$("#lrgawidget_sb_"+name+" .description-header").html(raw['total']+appnd);
			drawSparkline("#lrgawidget_spline_"+name, raw['data'], color);
		});
	}
}

function drawMainGraph(){
	lrWidgetSettings({action : "getSessions"}).done(function (data, textStatus, jqXHR) {
		if (data.status == "done" && data.setup !== 1){
			plotData = data.plotdata;
			plotTotalData = data.totalsForAllResults;
			if (!selectedPname){selectedPname = "sessions";}
			drawGraph(plotData[selectedPname], selectedPname);
			drawMainGraphWidgets(plotTotalData);
			$("#lrgawidget_sb_"+selectedPname).addClass("selected");
		}
	});	
}




$(document).ready(function(){
	
	dateRange = {start : moment().subtract(29, 'days').format('YYYY-MM-DD'),  end : moment().format('YYYY-MM-DD')};

    $('#lrgawidget_reportrange').html(moment(dateRange.start).format('MMMM D, YYYY') + ' - ' + moment(dateRange.end).format('MMMM D, YYYY'));
	$("[data-lrgawidget-reset]").on('click', function () {
		if (confirm("All saved authentication data will be removed.\nDo you want to continue ?!") == true) {
			lrWidgetSettings({action : "settingsReset"}).done(function (data, textStatus, jqXHR) {
				if (data.status == "done"){
					$('#lrga-wizard').wizard('selectedItem', {step: 1});
					$("[data-lrgawidget-reset]").hide();
				}
			});	
		}
	});
	
	$("#lrgawidget_main a[data-toggle='tab']").on('shown.bs.tab', function (e) {
		$("#lrgawidget_sessions_chart_tooltip").remove();
		
		if (this.hash == "#lrgawidget_settings_tab"){
			if (!setup){
				$('#lrga-wizard').wizard('selectedItem', {step: "lrga-profile"});
				$("#lrga-wizard .steps li").removeClass("complete");
				$("[data-lrgawidget-reset]").show();
			}
	    }else if (this.hash == "#lrgawidget_sessions_tab"){
			drawMainGraph();
		}
		
	});
 

    $("#lrgawidget_daterange_label").on('click', function (e) {
      e.preventDefault();
	  $('#lrgawidget a[href="#lrgawidget_gopro_tab"]').tab('show');

	});	

	$("[data-lrgawidget-plot]").on('click', function (e) {
		e.preventDefault();
		selectedPname = $(this).data('lrgawidget-plot');
		$("[data-lrgawidget-plot]").removeClass("selected");
		drawGraph(plotData[selectedPname] , selectedPname);
		$(this).addClass("selected");	
	});

    $('body').on('click', '#lrgawidget_panel_hide', function (e) {
		var wstatevalue = "";
		if ($(this).is(":checked")){
			$("#lrgawidget").show();
			wstatevalue = "show";
		}else{
			$("#lrgawidget").hide();
			wstatevalue = "hide";
		}
		lrWidgetSettings({action : "hideShowWidget", wstate: wstatevalue}).done(function (data, textStatus, jqXHR) {});	
	});

	$(".wrap:eq(1)").children("h1:first").remove();
	$("#adv-settings fieldset").append('<label for="lrgawidget_panel_hide"><input id="lrgawidget_panel_hide" type="checkbox" checked="checked">Lara, Google Analytics Dashboard Widget</label>');
	$("#lrgawidget_remove").on('click', function (e) {
		e.preventDefault(); 
		$("#lrgawidget_panel_hide").click();
	});		
	$(".daterangepicker").removeClass("daterangepicker dropdown-menu opensleft").addClass("lrga_bs daterangepicker custom-dropdown-menu opensleft");
	$('[data-toggle="lrgawidget_tooltip"]').tooltip();
	if (typeof actLrgaTabs !== 'undefined'){
		$("#lrgawidget a[href='#"+actLrgaTabs+"']").tab('show');
	}
    $(".lrgawidget_view_demo").colorbox({iframe:true, innerWidth:"80%", innerHeight:575, scrolling: false});
	
});


	
})(jQuery);