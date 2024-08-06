/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import * as d3 from "d3";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run_d3;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run_d3() {
  try {
    var width = 300;  //画布的宽度
    var height = 300;   //画布的高度
    var svg = d3.select("body")     //选择文档中的body元素
        .append("svg")          //添加一个svg元素
        .attr("width", width)       //设定宽度
        .attr("height", height);    //设定高度

    var dataset = [ 250 , 210 , 170 , 130 , 90 ];  //数据（表示矩形的宽度）
    var rectHeight = 25;   //每个矩形所占的像素高度(包括空白)
    svg.selectAll("rect")
        .data(dataset)
        .enter()
        .append("rect")
        .attr("x",20)
        .attr("y",function(d,i){
            return i * rectHeight;
        })
        .attr("width",function(d){
            return d;
        })
        .attr("height",rectHeight-2)
        .attr("fill","steelblue");

  } catch (error) {
    console.error(error);
  }
}