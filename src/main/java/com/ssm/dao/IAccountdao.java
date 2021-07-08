package com.ssm.dao;

import com.ssm.entity.Account;
import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Select;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository  //此注解代表这是一个持久层，用法类似@controller、@service
public interface IAccountdao {

    @Select("select * from account")
    public List<Account> findAll();
    @Insert("insert into account (name,money) value(#{name},#{money})")
    public void saveAccount(Account account);
}

