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
public partial class TestError1Category:IConfigCategory
{
	[JsonIgnore]
	public static TestError1Category Inst;


	private Dictionary<int, TestError1> dict = new Dictionary<int, TestError1>();
	[SerializeField]
	public List<TestError1> dataList= new List<TestError1>();

	public TestError1Category()
	{
		Inst = this;
	}


	public override void EndInit()
	{
		foreach (TestError1 config in dataList)
		{
			config.EndInit();
			this.dict.Add(config.Id,config);
		}
		this.AfterEndInit(); 
	}
	public TestError1 GetOne(int id)
	{
		if (dict.TryGetValue(id,out TestError1 c))
		{
			return c;
		}

		return null;
	}
	public bool isContain(int id)
	{
		return dict.ContainsKey(id);
	}
	public Dictionary<int, TestError1> GetDict()
	{
		return dict;
	}
}



[System.Serializable]
public partial class TestError1:IConfig
{
	public int Id; // 编号
	public string Name; // 名字
	public string BadType; // 坏类型
	public int[] ArrOnly1Col; // 单列数组
	public string Next; // 下一个
}

