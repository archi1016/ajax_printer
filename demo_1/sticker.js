
function hyAjax() {
	this.rq = new XMLHttpRequest();
	this.rq.owner = this;
	this.response_type = 'json';
	// this.csrf_token = Base.Element.get_meta_content('csrf-token');
	this.callback = {};
}
	hyAjax.prototype.append_property = function (pn, obj) {
		this.rq[pn] = obj;
	};
	hyAjax.prototype.set_response_type = function (t) {
		this.response_type = t;
	};
	hyAjax.prototype.set_callback = function (obj) {
		this.callback = obj;
	};
	hyAjax.prototype.launch_callback = function (name) {
		if (this.callback.hasOwnProperty(name)) {
			if (arguments.length > 1) {
				this.callback[name](arguments[1]);
			} else {
				this.callback[name]();
			}
		}
	};
	hyAjax.prototype.alert_error = function (message) {
		alert('Offline', message, 'OK');
	};

	// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	hyAjax.prototype.get = function (url) {
		var _self = this;
		var q = this.rq;

		this.launch_callback('show_processing');

		q.open('GET', url, true);
		q.onreadystatechange = _self.recv_status;
		q.onerror = function () {
			var obj = _self.rq.owner;

			obj.alert_error('error');
			obj.launch_callback('close_processing');
		};
		// q.setRequestHeader('X-CSRF-TOKEN', this.csrf_token);
		q.responseType = this.response_type;
		q.send();
	};
	hyAjax.prototype.post = function (url, post_data) {
		var _self = this;
		var q = this.rq;

		this.launch_callback('show_processing');

		q.open('POST', url, true);
		q.onreadystatechange = _self.recv_status;
		q.setRequestHeader('Content-type', 'application/x-www-form-urlencoded;charset=utf-8');
		// q.setRequestHeader('X-CSRF-TOKEN', this.csrf_token);
		q.responseType = this.response_type;
		q.send(post_data);
	};
	hyAjax.prototype.recv_status = function (ev) {
		var q = ev.target;
		var obj = q.owner;

		if (4 != q.readyState) return;
		if (0 == q.status) {
			obj.alert_error('net error');
			obj.launch_callback('close_processing');
			return;
		}
		if (200 != q.status) {
			obj.alert_error('return http status code: ' + q.status);
			obj.launch_callback('close_processing');
			return;
		}
		switch (obj.response_type) {
			case 'json':
				if (!q.response) {
					obj.alert_error('response format not JSON');
					obj.launch_callback('close_processing');
					return;
				}
				if('' == q.response.message) {
					obj.launch_callback('success', q.response);
				} else {
					obj.alert_error('#' + q.response.error + '<br />' + q.response.message);
					obj.launch_callback('error');
				}
				break;

			case 'document':
				if (!q.responseXML) {
					obj.alert_error('response format not XML');
					obj.launch_callback('close_processing');
					return;
				}
				obj.launch_callback('success', q.responseXML);
				break;

			case 'text':
				obj.launch_callback('success', q.responseText);
				break;

			case 'blob':
				if (!q.response) {
					obj.alert_error('response format not Blob');
					obj.launch_callback('close_processing');
					return;
				}
				obj.launch_callback('success', q.response);
				break;

			case 'arraybuffer':
				if (!q.response) {
					obj.alert_error('response format not ArrayBuffer');
					obj.launch_callback('close_processing');
					return;
				}
				obj.launch_callback('success', q.response);
				break;
		}
		obj.launch_callback('close_processing');
	};

// =============================================================================

var SoapSticker = {
	Parameter: {
		printer: 'Argox OS-214 plus series PPLA'
		,ajax_url: 'http://127.0.0.1:35427/?q'
		,xml_template: ''
	}

	// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	,Product: {
		Container: null
		,PrevLi: null
		,selected_index: -1
		,init: function (Owner) {
			this.Owner = Owner;
			this.Container = document.getElementById('PRODUCT-LIST');
		}
		,select_item: function (li) {
			this.Container.parentNode.className = this.Container.parentNode.className.replace(' focus', '');

			if (this.PrevLi) this.PrevLi.className = '';
			this.PrevLi = li;
			this.PrevLi.className = 'selected';
			this.selected_index = parseInt(li.getAttribute('item_index'));

			// this.Owner.Argument.Qty.value = '1';
			
			this.Owner.do_preview();
		}
		,render: function (Rows) {
			var _self = this;
			var tpl =	'<li>\
						<div class="no">{NO}.</div>\
						<div class="name">{NAME}</div>\
					 </li>\
					';
			var li;
			var i;
			var h;

			for (i=0; i<Rows.length; i++) {
				h = tpl;
				h = h.replace(/\{NO\}/g, i + 1);
				h = h.replace(/\{NAME\}/g, Rows[i].name);

				li = document.createElement('LI');
				li.setAttribute('item_index', i);
				li.innerHTML = h;

				li.onclick = function () {
					_self.select_item(this);
				};

				this.Container.appendChild(li);
			}
		}
	}

	// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	,Argument: {
		Weight: null
		,Qty: null
		,init: function (Owner) {
			this.Owner = Owner;

			this.Weight = document.getElementById('PRODUCT-WEIGHT');
			this.Qty = document.getElementById('STICKER-QTY');
		}
	}

	// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	,Preview: {
		Container: null
		,init: function (Owner) {
			this.Owner = Owner;
			this.Container = document.getElementById('PREVIEW-IMAGE');
		}
		,show: function (src) {
			if ('' != this.Container.src) URL.revokeObjectURL(this.Container.src);

			this.Container.src = src;
		}
		,draw: function (xml) {
			var _self = this;
			var aj = new hyAjax();

			aj.set_callback(
				{
					'success': function (bo) {
						_self.show(URL.createObjectURL(bo));
					}
				}
			);
			aj.set_response_type('blob');
			aj.post(this.Owner.get_ajax_url('preview'), xml);
		}
	}

	// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	,Printer: {
		init: function (Owner) {
			this.Owner = Owner;
		}
		,draw: function (xml) {
			var _self = this;
			var aj = new hyAjax();

			aj.set_callback(
				{
					'success': function () {
						
					}
				}
			);
			aj.post(this.Owner.get_ajax_url('print'), xml);
		}
	}

	// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	,DataRows: null
	,init: function (ps) {
		var _self = this;

		this.Parameter.xml_template = document.getElementById('XML-TEMPLATE').value;

		this.Product.init(this);
		this.Argument.init(this);
		this.Preview.init(this);
		this.Printer.init(this);
		this.DataRows = ps;

		this.Product.render(this.DataRows);

		document.getElementById('PREVIEW-STICKER').onclick = function () {
			if (_self.is_pass()) _self.do_preview();
		};
		document.getElementById('PRINT-STICKER').onclick = function () {
			if (_self.is_pass()) _self.do_print();
		};
	}
	,is_pass: function () {
		if (-1 == this.Product.selected_index) {
			this.Product.Container.parentNode.className += ' focus';
			alert('請選擇產品！');
			return false;
		}
		return true;
	}
	,do_preview: function () {
		this.Preview.draw(this.ret_script_xml());
	}
	,do_print: function () {
		this.Printer.draw(this.ret_script_xml());
	}
	,ret_script_xml: function () {
		var x = this.Parameter.xml_template;
		var p = this.DataRows[this.Product.selected_index];

		x = x.replace(/\{PRINTER\}/g, this.Parameter.printer);
		x = x.replace(/\{COPIES\}/g, this.Argument.Qty.value);

		x = x.replace(/\{NAME\}/g, p.name);
		x = x.replace(/\{BARCODE\}/g, p.barcode);
		x = x.replace(/\{PRICE\}/g, p.price);
		x = x.replace(/\{MATERIALS\}/g, p.materials);
		x = x.replace(/\{WEIGHT\}/g, this.Argument.Weight.value);
		return x;
	}
	,get_ajax_url: function (cmd) {
		return this.Parameter.ajax_url + '=' + cmd;
	}
};

window.addEventListener('load', function () {
	SoapSticker.init(ProductList);
});
