package dto;

import com.alibaba.excel.annotation.ExcelProperty;
import common.EnumValidAnnotation;
import enumeration.GenderEnum;
import lombok.Data;

import javax.validation.constraints.Digits;
import javax.validation.constraints.Max;
import javax.validation.constraints.NotNull;
import javax.validation.constraints.Size;

@Data
public class ResearcherImportData {
    @NotNull(message = "姓名不能为空")
    @Size(max = 16, message = "姓名过长")
    @ExcelProperty(value = "姓名", index = 0)
    private String name;

    @NotNull(message = "年龄不能为空")
    @Digits(integer = 2, fraction = 0, message = "年龄不是有效数字")
    @ExcelProperty(value = "年龄", index = 1)
    private Integer age;

    @NotNull(message = "性别不能为空")
    @EnumValidAnnotation(target = GenderEnum.class, message = "请选择有效性别")
    @ExcelProperty(value = "性别", index = 2)
    private String gender;
}
