using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using Newtonsoft.Json;
using UnityEngine;


/*
 *重点！！！！！！！！！！！！！！！！！！！！！！！
 *
 *该代码属于导表工具自动生成
 *切勿修改该代码 否则下次导表会被覆盖
 */


[System.Serializable]
[Config]
public partial class TestError2Category:IConfigCategory
{
	[JsonIgnore]
	public static TestError2Category Inst;


	private Dictionary<List<int>, TestError2> dict = new Dictionary<List<int>, TestError2>();
	[SerializeField]
	public List<TestError2> dataList= new List<TestError2>();

	public TestError2Category()
	{
		Inst = this;
	}


	public override void EndInit()
	{
		foreach (TestError2 config in dataList)
		{
			config.EndInit();
			this.dict.Add(config.Id,config);
		}
		this.AfterEndInit(); 
	}
	public TestError2 GetOne(List<int> id)
	{
		if (dict.TryGetValue(id,out TestError2 c))
		{
			return c;
		}

		return null;
	}
	public bool isContain(List<int> id)
	{
		return dict.ContainsKey(id);
	}
	public Dictionary<List<int>, TestError2> GetDict()
	{
		return dict;
	}
}



[System.Serializable]
public partial class TestError2:IConfig
{
	public List<int> Id; // 数组键
	public string Name; // 名字
	public int Name; // 重复
}

