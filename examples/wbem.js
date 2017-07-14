var ActiveX = require('../activex');
var conn = new ActiveX.Object('WbemScripting.SWbemLocator');
var svr = conn.ConnectServer('.', '\\root\\cimv2');
const resp = svr.ExecQuery('SELECT ProcessorId FROM Win32_Processor');
for (let i = 0; i < resp.Count; i += 1) {
	const properties = resp.ItemIndex(i).Properties_;
	let count = properties.Count;
	const propEnum = properties._NewEnum;
	while (count--) {
		const prop = propEnum.Next();
		if (prop) console.log(prop.Name + '=' + prop.Value);
	}
}