'use strict';

(function () {

	// The initialize function is run each time the page is loaded.
	Office.initialize = function (reason) {
		$(document).ready(function () {

			if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
				$('#msg').click(PrintMsg);
				$('#header').click(HeaderFormatting);
				$('#emplist').click(Displayemplist);

				$('#supportedVersion').html('This code is using Word 2016 or greater.');
			}
			else {
				$('#supportedVersion').html('This code requires Word 2016 or greater.');
			}

			Word.run(function (context) {

				const Paragraph = context.document.body;
				Paragraph.font.set({
					name: "Calibri",
					bold: false,
					size: 11,
					color: "Black",
				});
				context.document.body.clear();

				return context.sync();
			})
		});
	};

	function PrintMsg() {
		Word.run(function (context) {

			var thisDocument = context.document;

			var range = thisDocument.getSelection();
				//debugger;
			range.insertText('Employee List.\n', Word.InsertLocation.replace);


			return context.sync().then(function () {
				console.log('Added a quote from Ralph Waldo Emerson.');
			});

		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				alert('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}

	function Displayemplist() {
		Word.run(function (context) {
			const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
			var result = "hello";
			$.ajax({
				url: '../../api/Employee',
				type: 'GET',
				data: {
					empid: 0
				},
				contentType: 'application/json;charset=utf-8'
			}).done(function (data) {
				result = data;
				const tableData = [
					["Employee Name", "Department", "Joining Date", "Address", "Email", "Mobile No"]
				];
				for (var i = 0; i < result.length; i++) {
					tableData.push([result[i].employee_Name, result[i].department_Name, result[i].joiningDate, result[i].address, result[i].email, result[i].mobileNo]);
				}
				secondParagraph.insertTable(data.length + 1, 6, "", tableData);
				return context.sync();
			}).fail(function (status) {
				result = "Could not communicate with the server.";
			});
			
			return context.sync();
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				alert('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}

	

	function HeaderFormatting() {
		Word.run(function (context) {

			const secondParagraph = context.document.body.paragraphs.getFirst();
			secondParagraph.font.set({
				name: "Courier New",
				bold: true,
				size: 18,
				color: "Blue",
			});
			//secondParagraph.styleBuiltIn = Word.Style.intenseReference;
			//secondParagraph.style = "MyCustomStyle";

			return context.sync();
		})
			.catch(function (error) {
				console.log("Error: " + error);
				if (error instanceof OfficeExtension.Error) {
					console.log("Debug info: " + JSON.stringify(error.debugInfo));
				}
			});
	}
	
})();