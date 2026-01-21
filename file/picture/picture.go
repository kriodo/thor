package picture

func main() {
	//os.Remove("watermark.png")
	//// 打开原始图片
	//srcFile, err := os.Open("invoice_tmp.png")
	//if err != nil {
	//	panic(err)
	//}
	//defer srcFile.Close()
	//srcImg, _, err := image.Decode(srcFile)
	//if err != nil {
	//	panic(err)
	//}
	//
	//// 打开水印图片
	//watermarkFile, err := os.Open("watermark.png")
	//if err != nil {
	//	panic(err)
	//}
	//defer watermarkFile.Close()
	//watermarkImg, _, err := image.Decode(watermarkFile)
	//if err != nil {
	//	panic(err)
	//}
	//
	//// 调整水印图片大小（可选）
	//watermarkImg = resize.Resize(100, 100, watermarkImg, resize.Lanczos3) // 调整大小以适应需要的水印大小和位置
	//watermarkBounds := watermarkImg.Bounds()
	//watermarkWidth, watermarkHeight := watermarkBounds.Dx(), watermarkBounds.Dy()
	//
	//// 在原图上绘制水印（位置可以根据需要调整）
	//draw.Draw(srcImg, image.Rect(srcImg.Bounds().Max.X-watermarkWidth, srcImg.Bounds().Max.Y-watermarkHeight, srcImg.Bounds().Max.X, srcImg.Bounds().Max.Y), watermarkImg, image.Point{0, 0}, draw.Over)
	//
	//// 保存结果图片
	//outFile, err := os.Create("output.jpg") // 输出图片格式根据需要选择，如png、jpg等
	//if err != nil {
	//	panic(err)
	//}
	//defer outFile.Close()
	//err = jpeg.Encode(outFile, srcImg, &jpeg.Options{Quality: 95}) // 使用jpeg格式保存，根据需要选择其他格式和参数
	//if err != nil {
	//	panic(err)
	//}
}
