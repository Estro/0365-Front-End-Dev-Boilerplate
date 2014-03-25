// Martin Pomeroy - Architect 365 Helper Library

(function (APOINT) {
	"use strict";

	// GENERAL UTILITIES AND STUFF

	APOINT.utils = {
		cookies: {
			get: function (name) {
				name = name + "=";
				var ca = document.cookie.split(';');
				for (var i = 0; i < ca.length; i++) {
					var c = ca[i].trim();
					if (c.indexOf(name) === 0) return c.substring(name.length, c.length);
				}
				return false;
			},
			set: function (name, value, days) {
				var expires;
				if (days) {
					var date = new Date();
					date.setTime(date.getTime() + (days * 24 * 60 * 60 * 1000));
					expires = "; expires=" + date.toGMTString();
				} else {
					expires = "";
				}
				document.cookie = name + "=" + value + "; " + expires;
			},
			remove: function (name) {
				TESCO.Utils.cookies.set(name, "", -1);
			}
		},
		prettyDate: function (someDate) {
			var dd = someDate.getDate(),
				mm = someDate.getMonth(),
				monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
			return dd + ' ' + monthNames[mm];
		},
		logMessage: function (message) {
			console.log(message);
		},
		changeHexShade: function (hex, lum) {
			// validate hex string
			hex = String(hex).replace(/[^0-9a-f]/gi, '');
			if (hex.length < 6) {
				hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
			}
			lum = lum || 0;

			// convert to decimal and change luminosity
			var rgb = "#",
				c, i;
			for (i = 0; i < 3; i++) {
				c = parseInt(hex.substr(i * 2, 2), 16);
				c = Math.round(Math.min(Math.max(0, c + (c * lum)), 255)).toString(16);
				rgb += ("00" + c).substr(c.length);
			}

			if (rgb === "#000000") {
				rgb = "#4c4c4c";
			}
			return rgb;
		}
	};

	// UTILS TO HELP WITH MANGAGING SHAREPONT PAGES 

	APOINT.page = {
		inEditMode: function () {
			return SP.Ribbon.PageState.Handlers.isInEditMode();
		},
		getControlbyRow: function (value, elementType, number, visible) {
			var selector = visible ? elementType + ':visible' : elementType + ':hidden',
				row = $(".ms-standardheader:contains(" + value + ")").eq(0).parent().parent();
			return row.find(selector).eq(number - 1);
		}
	};

	// USER INFORMATION 

	APOINT.user = {
		getUserId: function () {
			// Returns logged in users id
			return _spPageContextInfo.userId;
		},
		getUserLoginName: function () {
			// Returns logged in users login name (email)
			return _spPageContextInfo.userLoginName;
		}
	};


	// URL HELPERS

	APOINT.url = {
		getQueryString: function (variable, query) {
			// Returns query string value from URL.
			// Can pass in a URL string via query parm
			if (query) {
				query = query.split('?')[1];
			} else {
				query = window.location.search.substring(1);
			}
			var vars = query.split("&");
			for (var i = 0; i < vars.length; i++) {
				var pair = vars[i].split("=");
				if (pair[0] == variable) {
					return unescape(pair[1]);
				}
			}
		},
		getSiteCollectionPath: function () {
			// Returns 'https://domain.sharepoint.com/teams/siteCollection
			return _spPageContextInfo.siteAbsoluteUrl;
		},
		getSitePath: function () {
			// Returns 'https://domain.sharepoint.com/teams/siteCollection/Site
			return _spPageContextInfo.webAbsoluteUrl;
		}
	};

	// INTERACT WITH SHAREPOINT DATA

	APOINT.data = {
		RenderViewFromArray: function (itemview, data) {
			var htmlout = '';
			$.each(data, function (i, value) {
				var cloneView = itemview;
				if (value.Icon) {
					value.Icon = value.Icon.Url;
				}
				if (value.Image) {
					value.Image = value.Image.Url;
				}

				if (value.File) {
					value.File = value.File.ServerRelativeUrl;
				}

				if(value.__metadata){
					value.__metadata = value.__metadata.media_src;
				}

				if (value.Display_x0020_Image) {
					value.Display_x0020_Image = value.Display_x0020_Image.Url;
				}

				$.each(value, function (key) {
					while (cloneView.indexOf('{' + key + '}') > 0) {
						cloneView = cloneView.replace('{' + key + '}', value[key]);
					}
				});
				htmlout += cloneView;
			});
			return htmlout;
		},
		updateListItem: function (spListURL, item, itemID, resultsFunction, library) {
			var meta = library ? "Item" : "ListItem";
			item.__metadata = {
				"type": "SP.Data." + spListURL.replace(/ /g, '_x0020_') + meta
			};
			spListURL = _spPageContextInfo.webServerRelativeUrl + '/_api/lists/getByTitle%28%27' + spListURL + '%27%29/items(' + itemID + ')';

			$.ajax({
				url: spListURL,
				type: "POST",
				contentType: "application/json;odata=verbose",
				data: JSON.stringify(item),
				headers: {
					"Accept": "application/json;odata=verbose",
						"X-RequestDigest": $("#__REQUESTDIGEST").val(),
						"X-HTTP-Method": "MERGE",
						"If-Match": "*"
				},
				success: function (data) {
					resultsFunction(data);
				},
				error: function (data) {
					APOINT.utils.logMessage("No Data returned from :: " + spListURL);
				}
			});
		},
		createListItem: function (spListURL, item, resultsFunction, site) {
			var siteURL = site ? _spPageContextInfo.webServerRelativeUrl + '/' + site : _spPageContextInfo.webServerRelativeUrl;
			item.__metadata = {
				"type": "SP.Data." + spListURL.replace(/ /g, '_x0020_') + "ListItem"
			};
			spListURL = siteURL + '/_api/lists/getByTitle%28%27' + spListURL + '%27%29/items';

			$.ajax({
				url: spListURL,
				type: "POST",
				contentType: "application/json;odata=verbose",
				data: JSON.stringify(item),
				headers: {
					"Accept": "application/json;odata=verbose",
						"X-RequestDigest": $("#__REQUESTDIGEST").val()
				},
				success: function (data) {
					resultsFunction(data);
				},
				error: function (data) {
					APOINT.utils.logMessage("No Data returned from :: " + spListURL);
				}
			});

		},
		getURL: function (spListURL, resultsFunction){	
			$.ajax({
					type: "GET",
					url: spListURL,
					dataType: 'json',
					headers: {
						"Accept": "application/json; odata=verbose"
					},
					success: function (data) {
					resultsFunction(data);
				},
				error: function () {
					APOINT.Utils.logMessage("Error getting data from list :: " + spListURL);
				}
			});
		},
		getListItems: function (spListURL, query, resultsFunction, number, site, doc, paged) {
			var siteURL = site ? _spPageContextInfo.webServerRelativeUrl + '/' + site : _spPageContextInfo.webServerRelativeUrl;
			if (!paged) {
				var isID = $.isNumeric(query);
				if (isID) {
					doc = doc ? '?$expand=File' : '';
					spListURL = siteURL + '/_api/lists/getByTitle%28%27' + spListURL + '%27%29/items(' + query + ')' + doc;

				} else {
					spListURL = spListURL.replace(/ /g, '%20');
					spListURL = siteURL + '/_api/lists/getByTitle%28%27' + spListURL + '%27%29/items?' + query;
					if (number) {
						spListURL += '&$top=' + number;
					} else {
						spListURL += '&$top=500';
					}
				}
			}

			$.ajax({
				type: "GET",
				url: spListURL,
				dataType: 'json',
				headers: {
					"Accept": "application/json; odata=verbose"
				},
				success: function (data) {
					if (data.d.results) {
						resultsFunction(data.d.results);
					} else if (data.d) {
						resultsFunction(data.d);
					} else {
						resultsFunction(data);
					}

					if (!number) {
						if (data.d.__next) {
							TESCO.Data.getListItems(data.d.__next, null, resultsFunction, false, true);
						}
					}
				},
				error: function () {
					APOINT.Utils.logMessage("Error getting data from list :: " + spListURL);
				}
			});
		},
		getListItemsSVC: function (spListURL, query, resultsFunction, number, site, paged) {
			var siteURL = site ? _spPageContextInfo.webServerRelativeUrl + '/' + site : _spPageContextInfo.webServerRelativeUrl;
			spListURL = siteURL + '/_layouts/_vti_bin/listData.svc/' + spListURL + '?' + query;
			if (number) {
				spListURL += '&$top=' + number;
			} else {
				spListURL += '&$top=500';
			}

			$.ajax({
				type: "GET",
				url: spListURL,
				dataType: 'json',
				headers: {
					"Accept": "application/json; odata=verbose"
				},
				success: function (data) {
					if (data.d.results) {
						resultsFunction(data.d.results);
					} else if (data.d) {
						resultsFunction(data.d);
					} else {
						resultsFunction(data);
					}

					if (!number) {
						if (data.d.__next) {
							TESCO.Data.getListItems(data.d.__next, null, resultsFunction, false, true);
						}
					}
				},
				error: function () {
					APOINT.Utils.logMessage("Error getting data from list :: " + spListURL);
				}
			});
		},
		uploadImage: function (library, imgContent, name, title, resultsFunction) {
			var site = _spPageContextInfo.webServerRelativeUrl;
			if (typeof (Uint8Array) === "function" || typeof (Uint8Array) === "object") {
				imgContent = APOINT.data.helpers.convertToBinary(imgContent);
				site = site.replace(/\s+/g, '');
				site = site.replace(',', '');

				if (site.length > 0) {
					site = site + '/';
				}

				var digest = $("#__REQUESTDIGEST").val(),
					url = _spPageContextInfo.webServerRelativeUrl + "/_api/web/lists/getByTitle(@TargetLibrary)/RootFolder/Files/add(url=@TargetFileName, overwrite='true')?@TargetLibrary='" + library + "'&@TargetFileName='" + name + "'&$select=ListItemAllFields&$expand=ListItemAllFields/Id";

				$.ajax({
					type: "POST",
					url: url,
					async: false,
					processData: false,
					data: imgContent,
					headers: {
						"accept": "application/json;odata=verbose",
							"content-type": "application/json;odata=verbose",
							"content-length": imgContent.length,
							"X-RequestDigest": $("#__REQUESTDIGEST").val()
					},
					success: function (data) {
						if (data.d.results) {
							resultsFunction(data.d.results);
						} else if (data.d) {
							resultsFunction(data.d);
						} else {
							APOINT.utils.logMessage("No Data returned from :: " + spListURL);
						}
					},
					error: function (err) {
						APOINT.utils.logMessage("No Data returned from :: " + spListURL);
					}
				});
			} else {
				var loadFrame = true;
				$('<iframe />').load(function () {
					if (loadFrame) {
						loadFrame = false;
						var imgData = $('#image-proxy').contents().find('.ief-imgInput'),
							siteInput = $('#image-proxy').contents().find('.ief-siteInput'),
							libraryInput = $('#image-proxy').contents().find('.ief-libraryInput'),
							nameInput = $('#image-proxy').contents().find('.ief-fileInput'),
							saveBtn = $('#image-proxy').contents().find('.ief-saveBtn');
							site = 'https://' + window.location.host + site;
							imgContent = imgContent.split('base64,')[1];

						if (imgData.length > 0) {
							imgData.val(imgContent);
							siteInput.val(site);
							libraryInput.val(library);
							nameInput.val(name);
							saveBtn.click();
							var counter = 0;
							var fieldValue;
							var timer = setInterval(function () {
								fieldValue = $('#image-proxy').contents().find('.ief-imgInput');
								if ($('#image-proxy').length > 0 && fieldValue.length > 0 && fieldValue.val().length > 0 && fieldValue.val().length < 10) {
									clearInterval(timer);
									resultsFunction(fieldValue.val());
								} else if (counter > 30) {
									clearInterval(timer);
								} else {
									fieldValue = null;
									counter++;
								}

							}, 2000);
						}
					}

				}).attr({
					src: '/teams/groupnews/Pages/image-proxy.aspx?isdlg=1',
					id: 'image-proxy'
				}).appendTo('footer');

			}
		},
		helpers: {
			convertToBinary: function (dataURI) {
				var BASE64_MARKER = ';base64,',
					base64Index = dataURI.indexOf(BASE64_MARKER) + BASE64_MARKER.length,
					base64 = dataURI.substring(base64Index),
					raw = window.atob(base64),
					rawLength = raw.length,
					array = new Uint8Array(new ArrayBuffer(rawLength)),
					i;
				for (i = 0; i < rawLength; i++) {
					array[i] = raw.charCodeAt(i);
				}
				return array;
			}
		},
		queries: {
			showFields: function (arry) {
				var query = '&$select=';
				$.each(arry, function (index, value) {
					query += value.replace(/ /g, '_x0020_');
					if (index !== arry.length - 1) {
						query += ',';
					}
				});
				return query;
			},
			orderBy: function (field, decending) {
				field = field.replace(/ /g, '_x0020_');
				var query = '&$orderby=' + field;
				if (decending) {
					query += '%20desc';
				}
				return query;
			},
			filterByDate: function (field, op, date) {
				field = field.replace(/ /g, '_x0020_');
				return '&$Filter=' + field + '%20' + op + '%20datetime%27' + date + '%27';
			},
			filterByField: function (field, value) {
				field = field.replace(/ /g, '_x0020_');
				return '&$filter=' + field + '%20eq%20%27' + value + '%27';
			},
			expandField: function (arry) {
				var query = '&$expand=';
				$.each(arry, function (index, value) {
					query += value.replace(/ /g, '_x0020_');
					if (index !== arry.length - 1) {
						query += ',';
					}
				});
				return query;
			}
		}

	};


})(window.APOINT = window.APOINT || {});

// -------------------------ISO POLYFILL -------------------------


if (typeof (Date.prototype.toISOString) != 'function') {
	Date.prototype.toISOString = function () {
		function pad(n) {
			return n < 10 ? '0' + n : n;
		}
		return (
		this.getUTCFullYear() + '-' + pad(this.getUTCMonth() + 1) + '-' + pad(this.getUTCDate()) + 'T' + pad(this.getUTCHours()) + ':' + pad(this.getUTCMinutes()) + ':' + pad(this.getUTCSeconds()) + '.' + pad(this.getUTCMilliseconds()) + 'Z');
	};
}