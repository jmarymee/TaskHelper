
module Ews {

    interface ISoap {
        toSoap(): string;
        getPrefix(): string;
        getPostfix(): string;
    }

    interface IRequest extends ISoap {

    }

    abstract class Request implements IRequest {

        getPrefix(): string {
            return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
    >

    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" soap:mustUnderstand="0" />
    </soap:Header>
    <soap:Body>
`;
        }

        getPostfix(): string {
            return `
    </soap:Body>
</soap:Envelope>
`;
        }

        toSoap(): string {
            throw "not implemented";
        }
    }

    abstract class Response {    }

    interface ITask extends ISoap {
    }

    export class Task implements ITask {
        getPrefix(): string {
            return `
<t:Task>
`;
        }

        getPostfix(): string {
            return `
</t:Task>
`;
        }

        constructor(private subject: string, private status?: string) {
        }

        // TODO: consider if we should have subject and status objects
        // which contain their own tags
        // Example: reuse this in FindItem with Restrictions, etc.
        toSoap(): string {
            var s: string = this.getPrefix();

            s += '<t:Subject>';
            s += this.subject;
            s += '</t:Subject>';

            if (this.status != null) {
                s += '<t:Status>';
                s += this.status;
                s += '</t:Status>';
            }

            s += this.getPostfix();
            return s;
        }
    }

    export class CreateTaskRequest extends Request {
        getPrefix(): string {
            return `
    <m:CreateItem MessageDisposition= "SaveOnly">
        <m:Items>
`;
        }

        getPostfix(): string {
            return `
        </m:Items>
    </m:CreateItem>
`;
        }

        constructor(private tasks: ITask[]) {
            super();
        }

        toSoap(): string {
            var envelope: string = '';

            envelope += super.getPrefix();
            envelope += this.getPrefix();

            for (var i = 0; i < this.tasks.length; i++) {
                envelope += this.tasks[i].toSoap();
            }

            envelope += this.getPostfix();
            envelope += super.getPostfix();
            envelope += '\n';

            return envelope;
        }

    }

    export interface IRestriction {
    }

    export class FindTaskRequest extends Request {

        getPrefix(): string {
            return `
<m:FindItem Traversal="Shallow">
    <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
    </m:ItemShape>
`;
        }

        getPostfix(): string {
            return `
    <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="tasks"/>
    </m:ParentFolderIds>
</m:FindItem>
`;
        }

        constructor(private restrictions?: IRestriction) {
            super();
        }

        toSoap(): string {
            var envelope: string = '';

            envelope += super.getPrefix();
            envelope += this.getPrefix();

            // TODO: add potential restrictions here

            envelope += this.getPostfix();
            envelope += super.getPostfix();

            return envelope;
        }
    }

}

// TODO: finish integrating these


//function getProperties() {
//    return '        <t:AdditionalProperties>' +
//        '          <t:FieldURI FieldURI="item:Subject" />' +
//    //        '          <t:FieldURI FieldURI="item:Priority" />' +
//        '          <t:FieldURI FieldURI="item:Status" />' +
//        '        </t:AdditionalProperties>';
//}

//function getAllTasks() {
//    var result = getBodyPrefix() +
//        '    <m:FindItem Traversal="Shallow">' +
//        '      <m:ItemShape>' +
//        '        <t:BaseShape>AllProperties</t:BaseShape>' +
//    //                    getProperties() +
//        '      </m:ItemShape>' +
//        getRestriction() +
//        '      <m:ParentFolderIds>' +
//        '        <t:DistinguishedFolderId Id="tasks"/>' +
//        '      </m:ParentFolderIds>' +
//        '    </m:FindItem>' +
//        getBodyPostfix();

//    return result;
//}

//function getRestriction() {
//    return '      <m:Restriction>' +
//    //        '        <t:And>' +
//        '          <t:IsNotEqualTo>' +
//        '            <t:FieldURI FieldURI="task:Status" />' +
//        '            <t:FieldURIOrConstant>' +
//        '              <t:Constant Value="2" />' +
//        '            </t:FieldURIOrConstant>' +
//        '          </t:IsNotEqualTo>' +
//    //        '        </t:And>' +
//        '      </m:Restriction>';
//}

//function addSort() {
//    return
//    '      <m:SortOrder>' +
//    '        <t:FieldOrder Order="Descending">' +
//    '          <t:FieldURI FieldURI="item:Priority" />' +
//    '        </t:FieldOrder>' +
//    '      </m:SortOrder>';
//}

//function getBodyPrefix() {
//    return '<?xml version="1.0" encoding="utf-8"?>' +
//        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
//        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
//        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
//        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
//        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
//        '  <soap:Header>' +
//        '    <t:RequestServerVersion Version="Exchange2013" soap:mustUnderstand="0" />' +
//        '  </soap:Header>' +
//        '  <soap:Body>';
//}

//function getBodyPostfix() {
//    return '  </soap:Body>' +
//        '</soap:Envelope>';
//}



//function getCreateTask() {
//    var result = getBodyPrefix() +
//        '    <m:CreateItem MessageDisposition="SaveOnly">' +
//        '      <m:Items>' +
//        getTask("Test EWS TaskHelper", "NotStarted") +
//        '      </m:Items>' +
//        '    </m:CreateItem>' +
//        getBodyPostfix();

//    return result;
//}


////        <t:Task>' +
////          <t:Subject>Test EWS TaskHelper</t:Subject>' +
////          <t:DueDate>2006-10-26T21:32:52</t:DueDate>' +
////          <t:Status>NotStarted</t:Status>' +
////        </t:Task>' +

//function getTask(subject, status) {
//    return '        <t:Task>' +
//        '          <t:Subject>' + subject + '</t:Subject>' +
//        '          <t:Status>' + status + '</t:Status>' +
//        '        </t:Task>';
//}


//function getSubjectRequest(id) {
//    // Return a GetItem operation request for the subject of the specified item. 
//    var result = getBodyPrefix() +
//        '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
//        '      <ItemShape>' +
//        '        <t:BaseShape>IdOnly</t:BaseShape>' +
//        '        <t:AdditionalProperties>' +
//        '            <t:FieldURI FieldURI="item:Subject"/>' +
//        '        </t:AdditionalProperties>' +
//        '      </ItemShape>' +
//        '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
//        '    </GetItem>' +
//        '  </soap:Body>' +
//        getBodyPostfix();

//    return result;
//}
